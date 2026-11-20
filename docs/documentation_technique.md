# Documentation technique : collecte des taux OANDA en XOF

## Fonctionnement
Le notebook `notebooks/oanda_taux_xof.ipynb` récupère, par l'API OANDA (`https://api.corporatefxservices.com`), le taux journalier Bid (champ `average_bid` de `/v2/rates/candles`) de MRU, GHS, USD, NGN et GNF contre XOF, du 1er janvier à la fin du mois visé, et écrit un fichier `<DEVISE>_XOF <Mois> <Année>.xlsx` par devise dans `Cours devises`.

Exécution recommandée : le 1er de chaque mois, pour le mois précédent (`MOIS_CIBLE = None`). Le notebook vérifie la date : si le mois n'est pas terminé, il n'appelle pas l'API. Les fichiers déjà présents ne sont ni retéléchargés ni affichés.

## Configuration de la clé
- Variable d'environnement `OANDA_FX_API_KEY` (recommandé pour l'automatisation), ou coffre Windows (`keyring`, proposé à la première saisie), ou saisie masquée.
- La clé n'est jamais écrite dans un fichier. Elle a été échangée en clair dans une conversation : la renouveler auprès d'OANDA.

## Paramètres (cellule 1)
`DOSSIER` (ou variable `COURS_DEVISES_DIR`), `DEVISES`, `MOIS_CIBLE` ("AAAA-MM"), `FORCER` (régénérer avec sauvegarde), `DATA_SET` (`None` ou `"BCEAO"`), `LANGUE_ENTETE` ("EN" ou "FR"), `ECRIRE_SI_INCOMPLET`, `SEUIL_SAUT`, `OUTILS` (test de connexion, spot, forwards).

## Format Excel reproduit
Feuille `rates_data` ; lignes 1 à 3 : en-tête ; lignes 4 et 5 vides ; ligne 6 : titres (`Period Start Date`, `XXX/XOF`) ; données dès la ligne 7, dates décroissantes (format `mm-dd-yy`), jours calendaires week-ends compris ; ligne vide ; 3 lignes de synthèse (moyenne, creux, pic) ; police Aptos Narrow 11 ; colonnes A et B de largeur 13 ; taux à 6 chiffres significatifs.

## Sécurité des données
Un fichier existant n'est remplacé (`FORCER = True`) qu'après copie dans `Cours devises/_sauvegardes`. Chaque exécution ajoute une ligne par devise à `_journal_collecte.csv` (écrit, refusé, échec, hors plage).

## Contrôles qualité (avant écriture)
Valeurs vides ou non numériques, taux nul ou négatif, jours manquants entre le 1er janvier et la fin du mois (fichier refusé par défaut : aucune valeur inventée), variations journalières supérieures à 15 % (alerte).

## Gestion d'erreurs
Réseau (3 nouvelles tentatives), 401/403 (clé ou droits), 404, 429 et 5xx (attente puis nouvel essai), fenêtre de dates non autorisée (arrêt propre, aucun fichier partiel), réponse vide. Une devise en échec n'interrompt pas les autres.

## Tests réalisés (API simulée, la clé réelle n'étant pas joignable depuis l'environnement de développement)
- Aller-retour sur `GHS_XOF Août 2026.xlsx` : 253 lignes régénérées, valeurs, dates, synthèse, format et police identiques. Seules 2 cellules d'en-tête diffèrent : le modèle contient `0,0%` et `hcc,corporatefxservices,com` (artefact de séparateur décimal) ; le notebook écrit `0.0%` et `hcc.corporatefxservices.com`.
- Scénarios d'erreur : clé refusée, fenêtre hors plage, réponse vide, 429 puis succès, réseau coupé, jours manquants (refus, puis écriture autorisée avec alerte), en-têtes FR.

## Limites et décisions à prendre
1. Fenêtre de dates : avec la clé actuelle, l'API n'accepte que 2025-05-13 à 2026-05-13. Les mois d'octobre 2026 ne peuvent pas être obtenus. Contacter OANDA (webservices@corporatefxservices.com) pour étendre la plage.
2. Équivalence `average_bid` (API) et « Bid » du convertisseur HCC : à valider en comparant, sur une période dans la fenêtre autorisée (par exemple Mars 2026), les valeurs de l'API avec les fichiers existants (`MOIS_CIBLE = "2026-03"`, `FORCER = True` sur une copie du dossier).
3. Jeu de données `BCEAO` (`DATA_SET`) : à confirmer si la référence BCEAO doit remplacer le jeu OANDA par défaut.
4. Langue des en-têtes : les modèles mélangent deux conventions ; `EN` par défaut.
5. Quotas : un lancement mensuel consomme 5 quotes.
