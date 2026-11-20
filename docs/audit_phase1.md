# Phase 1 : audit technique OANDA Exchange Rates API (Corporate FX Services)

Sources officielles lues le 2026-11-13 :
- Spécification OpenAPI 3.0.0, version 2.0.0 : https://api.corporatefxservices.com/openapi.json (copie dans docs/reference/)
- Portail développeurs : https://www.oanda.com/foreign-exchange-data-services/en/exchange-rates-api/api-developers/
- Plans : https://www.oanda.com/foreign-exchange-data-services/en/exchange-rates-api/api-plans/
- Dépôts GitHub OANDA : https://github.com/oanda

## Faits vérifiés
- Serveur : https://api.corporatefxservices.com
- Authentification : en-tête `Authorization: Bearer <clé>` (le paramètre `api_key` existe mais est déconseillé).
- Formats : json, xml, csv (extension du chemin, ex. `/v2/rates/spot.json`).
- Quota : chaque devise de cotation renvoyée compte pour un quote ; en-tête `X-Rate-Limit-Remaining` ; endpoint `/v2/remaining_quotes`.
- Pas de SDK Python officiel. Les SDK Exchange Rates (Node, Java, C#) sont archivés, avec renvoi vers la génération de client depuis l'OpenAPI.
- Les bindings `v20-python` / `oandapyV20` concernent l'API de trading : produit distinct.

## Endpoints
v1 : GET /v1/rates/{base}.{ext}
v2 : remaining_quotes, currencies, datasets/{data_set}, supported_forwards, rates/spot, rates/candle, rates/candles, rates/aggregated, rates/forward, rates/stream
v3 : instruments/{instrument}/orderBook, instruments/{instrument}/positionBook

## Plans (prix publics, USD)
Lite 450 / Premium 840 / Premium plus 1 160 / Premium + Crypto 1 680 par mois ; annuel 4 850 / 9 000 / 13 500 / 17 000. Essai gratuit de 7 jours, quotes illimités.
Les fonctions (forwards, streaming, spot 5 s, open/close, crypto) sont réparties par plan : le détail exact de la grille n'est pas lisible dans le HTML extrait, à confirmer.
