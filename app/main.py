"""Application de bureau pour extraire des taux de change via l'API OANDA Exchange Rates.

Lancement (depuis la racine du dépôt) :  python app/main.py
Dépendances : pandas, requests, openpyxl, matplotlib (Tkinter est fourni avec Python).
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk

import matplotlib
import pandas as pd

matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # noqa: E402
from matplotlib.figure import Figure  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from oanda_api import OandaClient, OandaError  # noqa: E402

# Charte Sunu : fond blanc, rouge et gris en aplats (texte blanc dessus), texte courant sombre.
# Les codes ci-dessous sont provisoires : à remplacer par les codes officiels de Sunu.
ROUGE = "#C8102E"
ROUGE_FONCE = "#9E0B23"
ROUGE_CLAIR = "#E3909D"
GRIS_FONCE = "#4B5058"
GRIS_FONCE_ACTIF = "#363A40"
GRIS_MOYEN = "#B4B9C0"
GRIS_CLAIR = "#EEF0F2"
GRIS_BORD = "#C9CED6"
FOND = "#FFFFFF"
PANNEAU = GRIS_CLAIR
TEXTE = "#1A1D21"
GRIS = "#4B5058"
POLICE = ("Segoe UI", 10)
POLICE_B = ("Segoe UI", 10, "bold")


def excel_ready(df: pd.DataFrame) -> pd.DataFrame:
    """Excel n'accepte pas les dates avec fuseau : conversion en UTC sans fuseau."""
    out = df.copy()
    for c in out.select_dtypes(include=["datetimetz"]).columns:
        out[c] = out[c].dt.tz_convert("UTC").dt.tz_localize(None)
        out = out.rename(columns={c: f"{c}_utc"})
    return out


def parse_date(texte: str, nom: str) -> date:
    try:
        return date.fromisoformat(texte.strip())
    except ValueError:
        raise ValueError(f"{nom} : format attendu AAAA-MM-JJ (reçu « {texte} »).") from None


class CurrencyPicker(ttk.Frame):
    """Liste de devises à sélection multiple avec filtre de recherche."""

    def __init__(self, master, titre: str, hauteur: int = 6):
        super().__init__(master)
        self.codes: list[str] = []
        self.libelles: dict[str, str] = {}
        self.selection: set[str] = set()
        self._visibles: list[str] = []
        ttk.Label(self, text=titre, style="Titre.TLabel").pack(anchor="w")
        self.filtre = tk.StringVar()
        self.filtre.trace_add("write", lambda *_: self._remplir())
        ttk.Entry(self, textvariable=self.filtre).pack(fill="x", pady=(2, 2))
        cadre = ttk.Frame(self)
        cadre.pack(fill="both", expand=True)
        self.liste = tk.Listbox(cadre, selectmode="multiple", exportselection=False, height=hauteur,
                                activestyle="none", relief="flat", highlightthickness=1,
                                highlightbackground=GRIS_BORD, selectbackground=ROUGE,
                                selectforeground="white", font=POLICE, bg="white", fg=TEXTE)
        asc = ttk.Scrollbar(cadre, orient="vertical", command=self.liste.yview)
        self.liste.configure(yscrollcommand=asc.set)
        self.liste.pack(side="left", fill="both", expand=True)
        asc.pack(side="right", fill="y")
        self.liste.bind("<<ListboxSelect>>", self._maj_selection)
        self.resume = ttk.Label(self, text="Aucune sélection", style="Info.TLabel")
        self.resume.pack(anchor="w", pady=(2, 0))

    def charger(self, codes: list[str], libelles: dict[str, str], defaut: list[str]):
        self.codes, self.libelles = codes, libelles
        self.selection = {c for c in defaut if c in codes}
        self.filtre.set("")
        self._remplir()

    def _remplir(self):
        f = self.filtre.get().strip().lower()
        self._visibles = [c for c in self.codes if f in c.lower() or f in self.libelles.get(c, "").lower()]
        self.liste.delete(0, "end")
        for i, c in enumerate(self._visibles):
            self.liste.insert("end", f"{c}   {self.libelles.get(c, '')}")
            if c in self.selection:
                self.liste.selection_set(i)
        premiere = next((i for i, c in enumerate(self._visibles) if c in self.selection), None)
        if premiere is not None:
            self.liste.see(premiere)
        self._resume()

    def _maj_selection(self, _evt=None):
        choisis = set(self.liste.curselection())
        for i, c in enumerate(self._visibles):
            (self.selection.add if i in choisis else self.selection.discard)(c)
        self._resume()

    def _resume(self):
        n = len(self.selection)
        self.resume.configure(text="Aucune sélection" if n == 0 else
                              ", ".join(sorted(self.selection)) if n <= 6 else f"{n} devises sélectionnées")

    def get(self) -> list[str]:
        return sorted(self.selection)


class DataTable(ttk.Frame):
    """Tableau avec barres de défilement."""

    def __init__(self, master):
        super().__init__(master)
        self.tree = ttk.Treeview(self, show="headings")
        v = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        h = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v.set, xscrollcommand=h.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        v.grid(row=0, column=1, sticky="ns")
        h.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

    def afficher(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        cols = list(df.columns)
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            large = 140 if str(df[c].dtype).startswith("datetime") else max(90, min(190, len(c) * 9))
            self.tree.column(c, width=large, anchor="w", stretch=False)
        txt = df.copy()
        for c in txt.select_dtypes(include=["datetimetz", "datetime"]).columns:
            txt[c] = txt[c].dt.strftime("%Y-%m-%d %H:%M")
        self.tree.tag_configure("alt", background=GRIS_CLAIR)
        for k, ligne in enumerate(txt.itertuples(index=False)):
            self.tree.insert("", "end", values=[("" if pd.isna(v) else v) for v in ligne],
                             tags=("alt",) if k % 2 else ())


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Taux de change OANDA")
        self.geometry("1360x860")
        self.minsize(1100, 680)
        self.configure(bg=FOND)
        self.client: OandaClient | None = None
        self.borne_min: date | None = None
        self.borne_max: date | None = None
        self.hist: pd.DataFrame | None = None
        self.spot_df: pd.DataFrame | None = None
        self.fwd_df: pd.DataFrame | None = None
        self.file: queue.Queue = queue.Queue()
        self._style()
        self._construire()
        self.after(100, self._boucle)
        env = os.environ.get("OANDA_FX_API_KEY", "")
        if env:
            self.var_cle.set(env)

    # ------------------------------------------------------------- style
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=FOND, foreground=TEXTE, font=POLICE, borderwidth=0)
        s.configure("TFrame", background=FOND)
        s.configure("Panneau.TFrame", background=PANNEAU)
        s.configure("TLabel", background=FOND)
        s.configure("Titre.TLabel", font=POLICE_B, foreground=TEXTE)
        s.configure("Info.TLabel", foreground=GRIS)
        s.configure("TLabelframe", background=FOND, bordercolor=GRIS_BORD, relief="solid", borderwidth=1)
        s.configure("TLabelframe.Label", background=GRIS_FONCE, foreground="white", font=POLICE_B, padding=(8, 2))
        s.configure("TButton", background=GRIS_FONCE, foreground="white", padding=(10, 5), relief="flat")
        s.map("TButton", background=[("active", GRIS_FONCE_ACTIF), ("disabled", GRIS_MOYEN)])
        s.configure("Accent.TButton", background=ROUGE, foreground="white", font=POLICE_B)
        s.map("Accent.TButton", background=[("active", ROUGE_FONCE), ("disabled", ROUGE_CLAIR)])
        s.configure("TEntry", fieldbackground="white", bordercolor=GRIS_BORD, padding=3)
        s.configure("TCombobox", fieldbackground="white", padding=3)
        s.configure("TNotebook", background=FOND, tabmargins=(0, 0, 0, 0))
        s.configure("TNotebook.Tab", padding=(16, 7), background=GRIS_CLAIR, foreground=TEXTE, font=POLICE)
        s.map("TNotebook.Tab", background=[("selected", GRIS_FONCE)], foreground=[("selected", "white")])
        s.configure("Treeview", background="white", fieldbackground="white", rowheight=22)
        s.configure("Treeview.Heading", background=GRIS_FONCE, foreground="white", font=POLICE_B, relief="flat")
        s.map("Treeview.Heading", background=[("active", GRIS_FONCE_ACTIF)])
        s.map("Treeview", background=[("selected", ROUGE)], foreground=[("selected", "white")])
        s.configure("TRadiobutton", background=FOND)
        s.configure("TCheckbutton", background=FOND)

    # ------------------------------------------------------ construction
    def _construire(self):
        bandeau = tk.Frame(self, bg=GRIS_FONCE, height=46)
        bandeau.pack(side="top", fill="x")
        bandeau.pack_propagate(False)
        tk.Label(bandeau, text="Taux de change OANDA", bg=GRIS_FONCE, fg="white",
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=16)
        tk.Label(bandeau, text="Extraction via l'API Exchange Rates", bg=GRIS_FONCE, fg="#D5D8DD",
                 font=POLICE).pack(side="left")
        gauche = ttk.Frame(self, padding=12)
        gauche.pack(side="left", fill="y")
        droite = ttk.Frame(self, padding=(0, 12, 12, 0))
        droite.pack(side="left", fill="both", expand=True)

        # Connexion
        cx = ttk.LabelFrame(gauche, text="Connexion", padding=10)
        cx.pack(fill="x")
        ttk.Label(cx, text="Clé API").pack(anchor="w")
        self.var_cle = tk.StringVar()
        self.entree_cle = ttk.Entry(cx, textvariable=self.var_cle, show="•", width=38)
        self.entree_cle.pack(fill="x", pady=(2, 2))
        self.var_voir = tk.BooleanVar()
        ttk.Checkbutton(cx, text="Afficher la clé", variable=self.var_voir,
                        command=lambda: self.entree_cle.configure(show="" if self.var_voir.get() else "•")).pack(anchor="w")
        self.btn_cx = ttk.Button(cx, text="Se connecter", style="Accent.TButton", command=self.connecter)
        self.btn_cx.pack(fill="x", pady=(6, 4))
        self.lbl_cx = ttk.Label(cx, text="Non connecté", style="Info.TLabel", wraplength=300, justify="left")
        self.lbl_cx.pack(anchor="w")

        # Devises
        dv = ttk.LabelFrame(gauche, text="Devises", padding=10)
        dv.pack(fill="both", expand=True, pady=(10, 0))
        self.sel_base = CurrencyPicker(dv, "Devise de base", hauteur=4)
        self.sel_base.pack(fill="both", expand=True)
        self.sel_quote = CurrencyPicker(dv, "Devise de cotation", hauteur=5)
        self.sel_quote.pack(fill="both", expand=True, pady=(8, 0))

        # Période
        pr = ttk.LabelFrame(gauche, text="Période", padding=10)
        pr.pack(fill="x", pady=(10, 0))
        ligne = ttk.Frame(pr)
        ligne.pack(fill="x")
        ttk.Label(ligne, text="Début").grid(row=0, column=0, sticky="w")
        ttk.Label(ligne, text="Fin").grid(row=0, column=1, sticky="w", padx=(10, 0))
        self.var_debut, self.var_fin = tk.StringVar(), tk.StringVar()
        ttk.Entry(ligne, textvariable=self.var_debut, width=14).grid(row=1, column=0)
        ttk.Entry(ligne, textvariable=self.var_fin, width=14).grid(row=1, column=1, padx=(10, 0))
        self.lbl_plage = ttk.Label(pr, text="Format AAAA-MM-JJ", style="Info.TLabel", wraplength=300, justify="left")
        self.lbl_plage.pack(anchor="w", pady=(4, 0))
        self.var_champs = tk.StringVar(value="all")
        ttk.Radiobutton(pr, text="Complet : open, close, moyenne, haut, bas", variable=self.var_champs, value="all").pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(pr, text="Moyenne, haut, bas seulement", variable=self.var_champs, value="partial").pack(anchor="w")

        # Onglets
        self.onglets = ttk.Notebook(droite)
        self.onglets.pack(fill="both", expand=True)
        self._onglet_historique()
        self._onglet_spot()
        self._onglet_forward()

        self.var_statut = tk.StringVar(value="Saisissez votre clé API puis cliquez sur « Se connecter ».")
        self.lbl_statut = ttk.Label(droite, textvariable=self.var_statut, style="Info.TLabel")
        self.lbl_statut.pack(anchor="w", pady=(6, 0))

    def _barre_export(self, parent, getter, nom):
        barre = ttk.Frame(parent)
        for fmt, lib in (("csv", "CSV"), ("json", "JSON"), ("xlsx", "Excel")):
            ttk.Button(barre, text=f"Exporter en {lib}",
                       command=lambda f=fmt: self.exporter(getter(), nom, f)).pack(side="left", padx=(0, 6))
        return barre

    def _onglet_historique(self):
        t = ttk.Frame(self.onglets, padding=10)
        self.onglets.add(t, text="Historique journalier")
        haut = ttk.Frame(t)
        haut.pack(fill="x")
        self.btn_hist = ttk.Button(haut, text="Extraire l'historique", style="Accent.TButton", command=self.extraire_historique)
        self.btn_hist.pack(side="left")
        self._barre_export(haut, lambda: self.hist, "historique").pack(side="right")
        self.lbl_hist = ttk.Label(t, text="", style="Info.TLabel")
        self.lbl_hist.pack(anchor="w", pady=(6, 0))
        pane = ttk.PanedWindow(t, orient="vertical")
        pane.pack(fill="both", expand=True, pady=(6, 0))
        bloc_g = ttk.Frame(pane)
        sel = ttk.Frame(bloc_g)
        sel.pack(fill="x")
        ttk.Label(sel, text="Paire affichée").pack(side="left")
        self.var_paire = tk.StringVar()
        self.cb_paire = ttk.Combobox(sel, textvariable=self.var_paire, state="readonly", width=14)
        self.cb_paire.pack(side="left", padx=8)
        self.cb_paire.bind("<<ComboboxSelected>>", lambda _e: self.tracer())
        self.fig = Figure(figsize=(8, 3.1), dpi=100, facecolor=FOND)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=bloc_g)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self._axe_vide()
        pane.add(bloc_g, weight=3)
        self.tab_hist = DataTable(pane)
        pane.add(self.tab_hist, weight=2)

    def _onglet_spot(self):
        t = ttk.Frame(self.onglets, padding=10)
        self.onglets.add(t, text="Taux spot")
        ligne = ttk.Frame(t)
        ligne.pack(fill="x")
        self.var_temps_reel = tk.BooleanVar(value=False)
        ttk.Checkbutton(ligne, text="Temps réel (selon abonnement)", variable=self.var_temps_reel).pack(side="left")
        ttk.Label(ligne, text="Date").pack(side="left", padx=(18, 4))
        self.var_spot_date = tk.StringVar()
        ttk.Entry(ligne, textvariable=self.var_spot_date, width=12).pack(side="left")
        ttk.Label(ligne, text="Heure UTC (pas de 15 min)").pack(side="left", padx=(12, 4))
        self.var_spot_heure = tk.StringVar(value="10:15")
        ttk.Entry(ligne, textvariable=self.var_spot_heure, width=7).pack(side="left")
        ligne2 = ttk.Frame(t)
        ligne2.pack(fill="x", pady=(8, 8))
        self.btn_spot = ttk.Button(ligne2, text="Récupérer les taux spot", style="Accent.TButton", command=self.recuperer_spot)
        self.btn_spot.pack(side="left")
        self._barre_export(ligne2, lambda: self.spot_df, "spot").pack(side="right")
        self.tab_spot = DataTable(t)
        self.tab_spot.pack(fill="both", expand=True)

    def _onglet_forward(self):
        t = ttk.Frame(self.onglets, padding=10)
        self.onglets.add(t, text="Taux forward")
        ttk.Label(t, text="Les taux forward dépendent de l'abonnement. Une erreur 403 signifie qu'ils ne sont pas inclus.",
                  style="Info.TLabel").pack(anchor="w")
        ligne = ttk.Frame(t)
        ligne.pack(fill="x", pady=(8, 8))
        ttk.Label(ligne, text="Base").pack(side="left")
        self.var_fb, self.var_fq = tk.StringVar(value="EUR"), tk.StringVar(value="USD")
        self.cb_fb = ttk.Combobox(ligne, textvariable=self.var_fb, width=8)
        self.cb_fb.pack(side="left", padx=(4, 14))
        ttk.Label(ligne, text="Cotation").pack(side="left")
        self.cb_fq = ttk.Combobox(ligne, textvariable=self.var_fq, width=8)
        self.cb_fq.pack(side="left", padx=(4, 14))
        ttk.Button(ligne, text="Lister les échéances", command=lambda: self.forward(False)).pack(side="left", padx=(0, 6))
        ttk.Button(ligne, text="Récupérer la courbe", style="Accent.TButton", command=lambda: self.forward(True)).pack(side="left")
        self._barre_export(ligne, lambda: self.fwd_df, "forward").pack(side="right")
        self.tab_fwd = DataTable(t)
        self.tab_fwd.pack(fill="both", expand=True)

    # ------------------------------------------------- exécution asynchrone
    def _boucle(self):
        try:
            while True:
                self.file.get_nowait()()
        except queue.Empty:
            pass
        self.after(100, self._boucle)

    def lancer(self, travail, succes, message="Requête en cours..."):
        self._statut(message)
        self.config(cursor="watch")

        def tache():
            try:
                res = travail()
            except Exception as e:  # noqa: BLE001
                self.file.put(lambda e=e: self._echec(e))
                return
            self.file.put(lambda: (self.config(cursor=""), succes(res)))

        threading.Thread(target=tache, daemon=True).start()

    def _statut(self, texte, erreur=False):
        self.var_statut.set(texte)
        self.lbl_statut.configure(foreground=TEXTE if erreur else GRIS, font=POLICE_B if erreur else POLICE)

    def _echec(self, e: Exception):
        self.config(cursor="")
        if isinstance(e, OandaError):
            if e.status in (401, 403) and "does not allow" not in e.message:
                texte = "Clé API refusée ou accès non autorisé. " + e.message
            elif e.status == 403:
                texte = "Non inclus dans votre abonnement. " + e.message
            else:
                texte = "L'API a refusé la requête. " + e.message
        else:
            texte = f"Connexion impossible : {e}"
        self._statut(texte, erreur=True)
        messagebox.showerror("Erreur", texte)

    # -------------------------------------------------------------- actions
    def connecter(self):
        cle = self.var_cle.get().strip()
        if not cle:
            messagebox.showwarning("Clé manquante", "Saisissez votre clé API.")
            return

        def travail():
            c = OandaClient(cle)
            return c, c.remaining_quotes(), c.currencies(), c.allowed_history_window()

        def succes(res):
            self.client, quotas, devises, (dmin, dmax) = res
            self.borne_min, self.borne_max = dmin, dmax
            codes = sorted(devises["code"])
            lib = dict(zip(devises["code"], devises["description"]))
            self.sel_base.charger(codes, lib, ["EUR"])
            self.sel_quote.charger(codes, lib, ["USD", "GBP", "CHF"])
            self.cb_fb["values"] = self.cb_fq["values"] = codes
            fin = dmax or date.today()
            debut = max(dmin, fin - timedelta(days=180)) if dmin else fin - timedelta(days=180)
            self.var_debut.set(debut.isoformat())
            self.var_fin.set(fin.isoformat())
            self.var_spot_date.set((fin - timedelta(days=1)).isoformat())
            self.lbl_plage.configure(text=f"Plage autorisée par votre clé : du {dmin} au {dmax}." if dmin
                                     else "Aucune limite de plage détectée. Format AAAA-MM-JJ.")
            self.lbl_cx.configure(foreground=TEXTE, text=f"Connecté. Quotes restants : {quotas['remaining_quotes']}, "
                                  f"utilisés : {quotas['used_quotes']}. {len(codes)} devises.")
            self._statut("Connecté. Choisissez les devises et la période, puis lancez l'extraction.")

        self.lancer(travail, succes, "Connexion à l'API...")

    def _verif_connexion(self) -> bool:
        if self.client is None:
            messagebox.showwarning("Non connecté", "Connectez-vous d'abord avec votre clé API.")
            return False
        return True

    def extraire_historique(self):
        if not self._verif_connexion():
            return
        bases, quotes = self.sel_base.get(), self.sel_quote.get()
        if not bases or not quotes:
            messagebox.showwarning("Devises", "Choisissez au moins une devise de base et une devise de cotation.")
            return
        try:
            debut, fin = parse_date(self.var_debut.get(), "Début"), parse_date(self.var_fin.get(), "Fin")
        except ValueError as e:
            messagebox.showwarning("Période", str(e))
            return
        if debut > fin:
            messagebox.showwarning("Période", "La date de début doit précéder la date de fin.")
            return
        # La date de fin de l'API semble exclusive : un jour de plus si la borne autorisée le permet.
        fin_req = fin + timedelta(days=1) if (self.borne_max is None or fin + timedelta(days=1) <= self.borne_max) else fin
        champs = ("all",) if self.var_champs.get() == "all" else ("averages", "highs", "lows")

        def travail():
            df, notes = self.client.candles(bases, quotes, debut, fin_req, fields=champs)
            if not df.empty:
                df = df[df["start_time"].dt.date <= fin].reset_index(drop=True)
            return df, notes

        def succes(res):
            df, notes = res
            self.hist = df
            if df.empty:
                self.lbl_hist.configure(text="Aucune donnée renvoyée pour cette sélection.")
                self.tab_hist.afficher(df)
                self._axe_vide()
                return
            n = df.groupby(["base_currency", "quote_currency"]).ngroups
            self.lbl_hist.configure(text=f"{len(df)} lignes, {df.shape[1]} colonnes, {n} paire(s)." +
                                    ("  " + " ".join(notes) if notes else ""))
            paires = [f"{b}/{q}" for b, q in df[["base_currency", "quote_currency"]].drop_duplicates().itertuples(index=False)]
            self.cb_paire["values"] = paires
            self.var_paire.set(paires[0])
            self.tab_hist.afficher(df)
            self.tracer()
            self._statut(f"Extraction terminée : {len(df)} lignes.")

        self.lancer(travail, succes, "Extraction en cours...")

    def _axe_vide(self):
        self.ax.clear()
        self.ax.set_facecolor("white")
        self.ax.text(0.5, 0.5, "Lancez une extraction pour afficher le graphique", ha="center", va="center",
                     color=GRIS, transform=self.ax.transAxes)
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.canvas.draw_idle()

    def tracer(self):
        if self.hist is None or self.hist.empty or not self.var_paire.get():
            return
        b, q = self.var_paire.get().split("/")
        d = self.hist[(self.hist.base_currency == b) & (self.hist.quote_currency == q)]
        x = d["start_time"].dt.tz_convert("UTC").dt.tz_localize(None)
        self.ax.clear()
        self.ax.set_facecolor("white")
        if {"low_midpoint", "high_midpoint"} <= set(d.columns):
            self.ax.fill_between(x, d["low_midpoint"], d["high_midpoint"], color=GRIS_MOYEN, alpha=0.55, linewidth=0, label="bas / haut")
        self.ax.plot(x, d["average_midpoint"], color=ROUGE, linewidth=1.6, label="midpoint moyen")
        self.ax.set_title(f"{b}/{q}", loc="left", fontsize=11, color=TEXTE)
        self.ax.set_ylabel("Taux")
        self.ax.grid(color=GRIS_BORD, alpha=0.6)
        self.ax.legend(loc="upper right", frameon=False)
        for s in ("top", "right"):
            self.ax.spines[s].set_visible(False)
        self.fig.autofmt_xdate()
        self.fig.tight_layout()
        self.canvas.draw_idle()

    def recuperer_spot(self):
        if not self._verif_connexion():
            return
        bases, quotes = self.sel_base.get(), self.sel_quote.get()
        if not bases or not quotes:
            messagebox.showwarning("Devises", "Choisissez au moins une devise de base et une devise de cotation.")
            return
        dt_iso = None
        if not self.var_temps_reel.get():
            try:
                d = parse_date(self.var_spot_date.get(), "Date")
                h = datetime.strptime(self.var_spot_heure.get().strip(), "%H:%M")
            except ValueError as e:
                messagebox.showwarning("Date", str(e) if "format attendu" in str(e) else "Heure : format attendu HH:MM.")
                return
            if h.minute % 15:
                messagebox.showwarning("Heure", "L'heure doit tomber sur un pas de 15 minutes (00, 15, 30, 45).")
                return
            dt_iso = datetime.combine(d, h.time()).strftime("%Y-%m-%dT%H:%M:00+00:00")

        def succes(df):
            self.spot_df = df
            self.tab_spot.afficher(df)
            self._statut(f"{len(df)} taux spot récupérés.")

        self.lancer(lambda: self.client.spot(bases, quotes, date_time=dt_iso), succes)

    def forward(self, courbe: bool):
        if not self._verif_connexion():
            return
        b, q = self.var_fb.get().strip().upper(), self.var_fq.get().strip().upper()
        if not b or not q:
            messagebox.showwarning("Devises", "Indiquez une devise de base et une devise de cotation.")
            return

        def succes(df):
            self.fwd_df = df
            self.tab_fwd.afficher(df)
            self._statut(f"{len(df)} lignes ({'courbe forward' if courbe else 'échéances disponibles'}).")

        self.lancer((lambda: self.client.forward(b, q)) if courbe else (lambda: self.client.supported_forwards(b, q)), succes)

    # --------------------------------------------------------------- export
    def exporter(self, df: pd.DataFrame | None, nom: str, fmt: str):
        if df is None or df.empty:
            messagebox.showinfo("Export", "Aucune donnée à exporter pour le moment.")
            return
        chemin = filedialog.asksaveasfilename(
            defaultextension=f".{fmt}", initialfile=f"{nom}.{fmt}",
            filetypes=[({"csv": "CSV", "json": "JSON", "xlsx": "Excel"}[fmt], f"*.{fmt}")])
        if not chemin:
            return
        try:
            if fmt == "csv":
                df.to_csv(chemin, index=False)
            elif fmt == "json":
                df.to_json(chemin, orient="records", date_format="iso", indent=2, force_ascii=False)
            else:
                excel_ready(df).to_excel(chemin, index=False)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Export", f"Échec de l'export : {e}")
            return
        self._statut(f"Fichier enregistré : {chemin}")


if __name__ == "__main__":
    App().mainloop()
