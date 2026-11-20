# Taux de change OANDA : application de bureau et notebook

Outils pour extraire des taux de change via l'API OANDA Exchange Rates
(serveur `api.corporatefxservices.com`, spécification OpenAPI 2.0.0 dans `docs/reference/`).

## Application de bureau (Windows, Mac, Linux)

Une fenêtre avec les champs de saisie : clé API, devises de base et de cotation (recherche intégrée),
période, puis trois onglets : historique journalier (graphique et tableau), taux spot, taux forward.
Export en CSV, JSON et Excel.

Installation (Python 3.11 ou plus récent) :

```bash
pip install -r requirements.txt
python app/main.py
```

Sous Windows, double-cliquez sur `lancer_appli.bat`.

Utilisation :
1. Collez votre clé API dans le champ « Clé API » (elle n'est enregistrée nulle part), puis cliquez sur « Se connecter ».
2. L'application affiche vos quotes restants et la plage d'historique autorisée par votre abonnement.
3. Choisissez les devises et la période, puis cliquez sur « Extraire l'historique ».

La clé peut aussi être préremplie avec la variable d'environnement `OANDA_FX_API_KEY`.

### Créer un exécutable Windows (optionnel)

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name TauxOANDA app/main.py
```

Le fichier `dist/TauxOANDA.exe` se lance sans installer Python.

## Notebook

`notebooks/oanda_extraction.ipynb` fait la même extraction sous forme de cellules à exécuter.

## Limites constatées avec une clé d'essai de 7 jours

- Taux en temps réel et forwards : refusés (HTTP 403), selon l'abonnement.
- Historique : plage limitée (une année glissante décalée de six mois environ). L'application la lit dans la réponse de l'API.
- Les dates doivent être au format complet AAAA-MM-JJThh:mm:ss+00:00, sur un pas de 15 minutes. L'application s'en charge.
- Les données sont journalières. L'intrajournalier passe par le taux spot à un instant précis.

## Charte graphique

Fond blanc, gris foncé et rouge en aplats. Les codes couleur sont définis en haut de `app/main.py`
et sont provisoires : à remplacer par les codes officiels de Sunu.

## Sécurité

Ne commitez jamais une clé API. Si elle a été partagée, régénérez-la dans le portail Self-Serve OANDA.
