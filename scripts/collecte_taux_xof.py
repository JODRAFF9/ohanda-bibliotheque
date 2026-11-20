"""Collecte des taux de change OANDA en XOF (MRU, GHS, USD, NGN, GNF).

Un fichier Excel par devise et par mois dans le dossier « Cours devises », puis export des moyennes
cumulees (1er janvier a fin de mois) dans « Cours devises/export ».

Utilisation :  python collecte_taux_xof.py                 (mois precedent)
               python collecte_taux_xof.py --mois 2026-01
               python collecte_taux_xof.py --mois 2026-01 --forcer --langue FR
La cle API est lue dans OANDA_FX_API_KEY, sinon le coffre Windows, sinon demandee (masquee).
Genere a partir du notebook oanda_taux_xof.ipynb : meme logique, memes controles.
"""
import argparse, subprocess, sys, importlib.util
from pathlib import Path

for mod, pip in {"requests": "requests", "pandas": "pandas", "openpyxl": "openpyxl", "keyring": "keyring"}.items():
    if importlib.util.find_spec(mod) is None and not getattr(sys, "frozen", False):
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip])

parser = argparse.ArgumentParser(description="Collecte des taux OANDA en XOF")
parser.add_argument("--mois", help='mois vise "AAAA-MM" (defaut : mois precedent)')
parser.add_argument("--forcer", action="store_true", help="regenerer un fichier deja present (ancien sauvegarde)")
parser.add_argument("--langue", choices=["EN", "FR"], default="EN", help="langue des en-tetes Excel")
parser.add_argument("--dossier", help="dossier Cours devises (defaut : ../Cours devises)")
ARGS = parser.parse_args()
BASE = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent

import os, getpass, shutil
from datetime import date, timedelta
from pathlib import Path

# ---------- À RÉGLER ----------
DOSSIER = Path(ARGS.dossier or os.environ.get("COURS_DEVISES_DIR") or (BASE.parent / "Cours devises"))
MOIS_CIBLE = ARGS.mois
FORCER = ARGS.forcer
LANGUE_ENTETE = ARGS.langue

# ---------- RÉGLAGES AVANCÉS (rarement modifiés) ----------
DEVISES = ["MRU", "GHS", "USD", "NGN", "GNF"]
CIBLE = "XOF"
TELECHARGER = True            # False : ne rien télécharger
DATA_SET = None               # None = jeu OANDA par défaut ; "BCEAO" = jeu BCEAO (à confirmer avec OANDA)
ECRIRE_SI_INCOMPLET = False   # False : un mois avec des jours manquants n'est pas écrit
SEUIL_SAUT = 0.15             # alerte si le taux varie de plus de 15 % d'un jour à l'autre
DOSSIER_SAUVEGARDE = DOSSIER / "_sauvegardes"
JOURNAL = DOSSIER / "_journal_collecte.csv"

MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
           "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

ENTETES = {
    "EN": ["Daily BID rates @ +/- 0.0%", "OANDA Exchange Rates", "https://hcc.corporatefxservices.com",
           "Period Start Date", ["Period Average", "Period Low", "Period High"]],
    "FR": ["Taux SOUMISSION quotidiens @ +/- 0.0%", "Taux FX OANDA", "https://fxds-hcc.oanda.com",
           "Date de début de la période", ["Moyenne de la période", "Creux de la période", "Pic de la période"]],
}

def mois_vise():
    """Mois précédent (ou MOIS_CIBLE) : renvoie (1er janvier, dernier jour du mois)."""
    if MOIS_CIBLE:
        a, m = map(int, MOIS_CIBLE.split("-"))
        premier = date(a, m, 1)
    else:
        premier = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
    fin = (premier.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    return date(fin.year, 1, 1), fin

import requests, time, re
import pandas as pd
import keyring

API_URL = "https://api.corporatefxservices.com"


class ErreurAPI(RuntimeError):
    pass


class FenetreHorsPlage(ErreurAPI):
    def __init__(self, debut, fin):
        super().__init__(f"dates autorisées par la clé : {debut} au {fin}")
        self.debut, self.fin = debut, fin


def cle_api():
    cle = os.environ.get("OANDA_FX_API_KEY") or keyring.get_password("oanda_fx_api", "cle")
    if cle:
        return cle.strip()
    cle = getpass.getpass("Clé API OANDA : ").strip()
    if input("Mémoriser dans le coffre Windows ? (o/n) ").strip().lower().startswith("o"):
        keyring.set_password("oanda_fx_api", "cle", cle)
    return cle


def api_get(session, chemin, tentatives=4, **params):
    params = {k: v for k, v in params.items() if v is not None}
    for t in range(tentatives):
        try:
            r = session.get(f"{API_URL}{chemin}.json", params=params, timeout=30)
        except requests.RequestException as e:
            if t == tentatives - 1:
                raise ErreurAPI(f"connexion impossible : {e}")
            time.sleep(2 ** t)
            continue
        if r.status_code == 429 or r.status_code >= 500:
            if t == tentatives - 1:
                raise ErreurAPI(f"HTTP {r.status_code} persistant (limite ou service indisponible)")
            time.sleep(int(r.headers.get("Retry-After", 2 ** (t + 1))))
            continue
        if r.status_code in (401, 403):
            raise ErreurAPI(f"HTTP {r.status_code} : clé refusée ou fonction non incluse dans l'abonnement ({r.text[:200]})")
        if r.status_code == 404:
            raise ErreurAPI(f"HTTP 404 : donnée ou paire indisponible ({r.text[:200]})")
        if not r.ok:
            m = re.search(r">= (\d{4}-\d{2}-\d{2}) and <= (\d{4}-\d{2}-\d{2})", r.text)
            if m:
                raise FenetreHorsPlage(m.group(1), m.group(2))
            raise ErreurAPI(f"HTTP {r.status_code} sur {chemin} : {r.text[:300]}")
        restant = r.headers.get("X-Rate-Limit-Remaining")
        if restant is not None:
            session.dernier_restant = restant
        return r.json()


def iso(jour):
    """Format exigé : AAAA-MM-JJThh:mm:ss+00:00 ; début et fin doivent avoir la même heure et minute (ici 00:00 UTC)."""
    return f"{jour.isoformat()}T00:00:00+00:00"


def taux_bid(session, base, debut, fin):
    """Taux journaliers Bid (average_bid) de base/XOF, du plus récent au plus ancien. Lève ErreurAPI si vide."""
    data = api_get(session, "/v2/rates/candles", base=base, quote=CIBLE, fields="averages",
                   start_time=iso(debut), end_time=iso(fin + timedelta(days=1)), data_set=DATA_SET)
    quotes = data.get("quotes") or []
    if not quotes:
        raise ErreurAPI(f"aucune donnée renvoyée pour {base}/{CIBLE}")
    df = pd.DataFrame(quotes)
    df["date"] = pd.to_datetime(df["start_time"], utc=True).dt.tz_localize(None).dt.normalize()
    df["taux"] = pd.to_numeric(df["average_bid"], errors="coerce")
    df = df[(df["date"].dt.date >= debut) & (df["date"].dt.date <= fin)]
    return (df[["date", "taux"]].drop_duplicates("date").sort_values("date", ascending=False)
              .reset_index(drop=True))

from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime


def controles(df, debut, fin):
    """Renvoie (erreurs bloquantes, avertissements)."""
    err, warn = [], []
    attendus = pd.date_range(debut, fin, freq="D")
    if df["taux"].isna().any():
        err.append("valeurs vides ou non numériques")
    if (df["taux"] <= 0).any():
        err.append("taux nul ou négatif")
    manq = attendus.difference(df["date"])
    if len(manq):
        liste = ", ".join(d.strftime("%d/%m") for d in manq[:8]) + ("..." if len(manq) > 8 else "")
        (warn if ECRIRE_SI_INCOMPLET else err).append(f"{len(manq)} jour(s) manquant(s) sur {len(attendus)} : {liste}")
    s = df.sort_values("date")["taux"]
    sauts = s.pct_change().abs()
    if (sauts > SEUIL_SAUT).any():
        n = int((sauts > SEUIL_SAUT).sum())
        warn.append(f"{n} variation(s) journalière(s) > {SEUIL_SAUT:.0%} à vérifier")
    return err, warn


def sauvegarder(chemin):
    if chemin.exists():
        DOSSIER_SAUVEGARDE.mkdir(parents=True, exist_ok=True)
        copie = DOSSIER_SAUVEGARDE / f"{chemin.stem}_{datetime.now():%Y%m%d_%H%M%S}{chemin.suffix}"
        shutil.copy2(chemin, copie)
        return copie


def ecrire_fichier(df, base, fin, langue=None):
    """Écrit <BASE>_XOF <Mois> <Année>.xlsx au format des modèles. Sauvegarde l'existant avant remplacement."""
    t1, t2, t3, titre, synth = ENTETES[langue or LANGUE_ENTETE]
    df = df.sort_values("date", ascending=False)
    taux = [float(f"{x:.6g}") for x in df["taux"]]
    wb = Workbook()
    ws = wb.active
    ws.title = "rates_data"
    police = Font(name="Aptos Narrow", size=11)
    for ligne in ([t1, None], [t2, None], [t3, None], [None, None], [None, None], [titre, f"{base}/{CIBLE}"]):
        ws.append(ligne)
    for d, t in zip(df["date"], taux):
        ws.append([d.to_pydatetime(), t])
        ws.cell(ws.max_row, 1).number_format = "mm-dd-yy"
    ws.append([None, None])
    vals = [float(f"{sum(taux) / len(taux):.6g}"), min(taux), max(taux)]
    for lib, v in zip(synth, vals):
        ws.append([lib, v])
    for row in ws.iter_rows():
        for c in row:
            c.font = police
    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 13
    chemin = DOSSIER / f"{base}_{CIBLE} {MOIS_FR[fin.month - 1]} {fin.year}.xlsx"
    DOSSIER.mkdir(parents=True, exist_ok=True)
    copie = sauvegarder(chemin)
    wb.save(chemin)
    return chemin, copie


def journaliser(ligne):
    ligne = {"horodatage": datetime.now().isoformat(timespec="seconds"), **ligne}
    pd.DataFrame([ligne]).to_csv(JOURNAL, mode="a", header=not JOURNAL.exists(), index=False, encoding="utf-8-sig")

debut, fin = mois_vise()
nouveaux = {}   # seules les données téléchargées pendant cette exécution sont affichées
nom_mois = f"{MOIS_FR[fin.month - 1]} {fin.year}"
print(f"Aujourd'hui : {date.today()} | mois visé : {nom_mois} (du {debut} au {fin})")


def collecter(session=None):
    global nouveaux
    nouveaux = {}
    bilan = []
    a_faire = []
    for d in DEVISES:
        existants = list(DOSSIER.glob(f"{d}_{CIBLE} {nom_mois}.*"))
        if existants and not FORCER:
            print("DÉJÀ PRÉSENT", existants[0].name); bilan.append((d, "déjà présent", ""))
        else:
            a_faire.append(d)
    if not a_faire:
        print("Rien à télécharger : le mois est complet dans le dossier.")
        return
    session = session or requests.Session()
    if "Authorization" not in session.headers:
        session.headers["Authorization"] = f"Bearer {cle_api()}"
    try:
        print("Quotes restants :", api_get(session, "/v2/remaining_quotes"))
    except ErreurAPI as e:
        print("ARRÊT :", e)
        return
    for d in a_faire:
        try:
            df = taux_bid(session, d, debut, fin)
            err, warn = controles(df, debut, fin)
            for w in warn:
                print(f"  ATTENTION {d} : {w}")
            if err:
                print(f"{d} -> NON ÉCRIT : " + " ; ".join(err)); bilan.append((d, "non écrit", " ; ".join(err)))
                journaliser({"devise": d, "mois": nom_mois, "statut": "refusé", "detail": " ; ".join(err)})
                continue
            chemin, copie = ecrire_fichier(df, d, fin)
            nouveaux[d] = df.sort_values("date")
            print(d, "->", chemin.name, f"({len(df)} jours)" + (f", ancien fichier sauvegardé : {copie.name}" if copie else ""))
            bilan.append((d, "écrit", f"{len(df)} jours")); journaliser({"devise": d, "mois": nom_mois, "statut": "écrit", "detail": f"{len(df)} jours; {' ; '.join(warn)}"})
        except FenetreHorsPlage as e:
            print(f"ARRÊT : l'API n'autorise que {e.debut} au {e.fin} avec cette clé ; période demandée {debut} au {fin}.")
            bilan.append((d, "hors plage", f"{e.debut} au {e.fin}"))
            print("Aucun fichier écrit. Pour étendre la plage : webservices@corporatefxservices.com.")
            journaliser({"devise": d, "mois": nom_mois, "statut": "hors plage", "detail": f"{e.debut}..{e.fin}"})
            break
        except ErreurAPI as e:
            print(d, "-> ÉCHEC :", e); bilan.append((d, "échec", str(e)[:80]))
            journaliser({"devise": d, "mois": nom_mois, "statut": "échec", "detail": str(e)[:200]})
    print("\nRécapitulatif :")
    print(pd.DataFrame(bilan, columns=["Devise", "Résultat", "Détail"]).to_string(index=False))


if not TELECHARGER:
    print("Téléchargement désactivé (TELECHARGER = False).")
elif fin >= date.today():
    print("Impossible pour l'instant : le mois n'est pas terminé. Réessayez à partir du", fin + timedelta(days=1))
else:
    collecter()

import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

LIBELLES_MOYENNE = ("Period Average", "Moyenne de la période")
DOSSIER_EXPORT = DOSSIER / "export"
FICHIER_MOYENNES = DOSSIER_EXPORT / "moyennes_cumulees_XOF.xlsx"


def moyenne_du_fichier(chemin):
    """Valeur de la ligne de synthèse « moyenne » d'un fichier mensuel (None si introuvable)."""
    ws = load_workbook(chemin, read_only=True, data_only=True).active
    for a, b in ws.iter_rows(min_row=1, max_col=2, values_only=True):
        if isinstance(a, str) and a.strip().startswith(LIBELLES_MOYENNE):
            return b
    return None


def exporter_moyennes():
    """Un onglet « Mensuel » : mois en lignes, devises en colonnes, moyenne du 1er janvier à la fin du mois."""
    valeurs = {}
    for f in DOSSIER.glob(f"*_{CIBLE} *.xlsx"):
        m = re.match(rf"^([A-Z]{{3}})_{CIBLE}\s+(\S+)\s+(\d{{4}})\.xlsx$", f.name)
        if not m or m.group(2) not in MOIS_FR:
            continue
        dev, mois, an = m.group(1), MOIS_FR.index(m.group(2)) + 1, int(m.group(3))
        if dev in DEVISES:
            valeurs[(an, mois, dev)] = moyenne_du_fichier(f)
    if not valeurs:
        print("Aucun fichier mensuel trouvé : export non créé.")
        return None
    periodes = sorted({(a, m) for a, m, _ in valeurs})
    wb = Workbook()
    ws = wb.active
    ws.title = "Mensuel"
    ws.append(["Mois"] + [f"{d}/{CIBLE}" for d in DEVISES])
    for c in ws[1]:
        c.font = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="D0021B")
        c.alignment = Alignment(horizontal="center")
    for a, m in periodes:
        ws.append([f"{MOIS_FR[m - 1]} {a}"] + [valeurs.get((a, m, d)) for d in DEVISES])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Aptos Narrow", size=11)
    ws.column_dimensions["A"].width = 18
    for col in "BCDEF"[:len(DEVISES)]:
        ws.column_dimensions[col].width = 13
    ws.freeze_panes = "B2"
    DOSSIER_EXPORT.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(FICHIER_MOYENNES)   # fichier dérivé, régénéré à chaque fois : pas de sauvegarde nécessaire
    except PermissionError:
        print("Export impossible : fermer moyennes_cumulees_XOF.xlsx dans Excel puis relancer cette cellule.")
        return None
    vides = sum(1 for p in periodes for d in DEVISES if valeurs.get((*p, d)) is None)
    print(f"export/{FICHIER_MOYENNES.name} : {len(periodes)} mois, {len(DEVISES)} devises" + (f", {vides} case(s) vide(s)" if vides else ""))
    return FICHIER_MOYENNES


exporter_moyennes()
