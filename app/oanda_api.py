"""Client minimal pour l'API OANDA Exchange Rates (spécification OpenAPI v2.0.0).

Serveur : https://api.corporatefxservices.com, authentification Bearer.
Seuls les endpoints documentés dans docs/reference/ sont utilisés.
"""
from __future__ import annotations

import re
import time
from datetime import date

import pandas as pd
import requests

BASE_URL = "https://api.corporatefxservices.com"
_WINDOW_RE = re.compile(r">= (\d{4}-\d{2}-\d{2}) and <= (\d{4}-\d{2}-\d{2})")


class OandaError(RuntimeError):
    """Erreur renvoyée par l'API (code HTTP et message explicite)."""

    def __init__(self, status: int, message: str):
        super().__init__(f"HTTP {status} : {message}")
        self.status = status
        self.message = message


def _list(x) -> list[str]:
    return [x] if isinstance(x, str) else list(x)


class OandaClient:
    def __init__(self, api_key: str, base_url: str = BASE_URL, timeout: int = 30, retries: int = 3):
        self.session = requests.Session()
        self.session.headers["Authorization"] = f"Bearer {api_key.strip()}"
        self.base_url, self.timeout, self.retries = base_url, timeout, retries
        self.remaining: str | None = None

    # --- bas niveau -------------------------------------------------------
    def get(self, path: str, **params) -> dict:
        """GET JSON. Réessaie sur erreur réseau, 429 et 5xx ; lève OandaError sinon."""
        params = {k: v for k, v in params.items() if v is not None}
        for attempt in range(self.retries):
            try:
                r = self.session.get(f"{self.base_url}{path}.json", params=params, timeout=self.timeout)
            except requests.RequestException:
                if attempt == self.retries - 1:
                    raise
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 429 or r.status_code >= 500:
                if attempt == self.retries - 1:
                    raise OandaError(r.status_code, r.text[:300])
                time.sleep(2 ** attempt)
                continue
            if not r.ok:
                try:
                    msg = r.json().get("message", r.text)
                except ValueError:
                    msg = r.text
                raise OandaError(r.status_code, msg[:400])
            self.remaining = r.headers.get("X-Rate-Limit-Remaining", self.remaining)
            return r.json()
        raise RuntimeError("inatteignable")

    # --- métadonnées ------------------------------------------------------
    def remaining_quotes(self) -> dict:
        return self.get("/v2/remaining_quotes")

    def currencies(self) -> pd.DataFrame:
        return pd.DataFrame(self.get("/v2/currencies")["currencies"])

    def allowed_history_window(self) -> tuple[date | None, date | None]:
        """Plage d'historique acceptée par la clé, lue dans le message d'erreur 400
        d'une requête volontairement hors plage. (None, None) si aucune limite détectée."""
        try:
            self.get("/v2/rates/candles", base="EUR", quote="USD",
                     start_time="2000-01-01T00:00:00+00:00", end_time="2000-01-02T00:00:00+00:00",
                     fields="averages")
        except OandaError as e:
            m = _WINDOW_RE.search(e.message)
            if m:
                return date.fromisoformat(m[1]), date.fromisoformat(m[2])
        return None, None

    # --- taux -------------------------------------------------------------
    def spot(self, base, quote=None, date_time: str | None = None, data_set: str | None = None) -> pd.DataFrame:
        data = self.get("/v2/rates/spot", base=_list(base), quote=_list(quote) if quote else None,
                        date_time=date_time, data_set=data_set)
        df = pd.DataFrame(data["quotes"])
        for c in ("bid", "ask", "midpoint"):
            df[c] = pd.to_numeric(df[c])
        df["date_time"] = pd.to_datetime(df["date_time"], utc=True)
        return df

    def candles(self, base, quote, start: date, end: date, fields=("all",), data_set: str | None = None):
        """Chandeliers journaliers. Renvoie (DataFrame, notes).

        La requête est découpée selon la limite documentée de 10 000 (jours x paires).
        Si l'abonnement refuse open/close (403), repli sur moyenne, haut et bas.
        """
        base, quote, notes = _list(base), _list(quote), []
        jours_max = max(1, min(10_000 // (len(base) * len(quote)), 365))
        fields, parts = list(fields), []
        cur, last = pd.Timestamp(start), pd.Timestamp(end)
        while cur <= last:
            stop = min(cur + pd.Timedelta(days=jours_max - 1), last)
            args = dict(base=base, quote=quote, data_set=data_set,
                        start_time=cur.strftime("%Y-%m-%dT00:00:00+00:00"),
                        end_time=stop.strftime("%Y-%m-%dT00:00:00+00:00"))
            try:
                data = self.get("/v2/rates/candles", fields=fields, **args)
            except OandaError as e:
                if e.status != 403 or "all" not in fields:
                    raise
                notes.append("Open et close non inclus dans l'abonnement : repli sur moyenne, haut et bas.")
                fields = ["averages", "highs", "lows"]
                data = self.get("/v2/rates/candles", fields=fields, **args)
            parts.append(pd.DataFrame(data["quotes"]))
            cur = stop + pd.Timedelta(days=1)
        df = pd.concat(parts, ignore_index=True)
        if df.empty:
            return df, notes
        df = df.drop_duplicates(subset=["base_currency", "quote_currency", "start_time"])
        for c in df.columns:
            if c.endswith(("_bid", "_ask", "_midpoint")):
                df[c] = pd.to_numeric(df[c])
        for c in ("start_time", "open_time", "close_time", "source_date"):
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], utc=True)
        df = df.sort_values(["base_currency", "quote_currency", "start_time"]).reset_index(drop=True)
        return df, list(dict.fromkeys(notes))

    # --- forwards ---------------------------------------------------------
    def supported_forwards(self, base, quote) -> pd.DataFrame:
        return pd.DataFrame(self.get("/v2/supported_forwards", base=_list(base), quote=_list(quote))["supported_forwards"])

    def forward(self, base: str, quote: str, tenor=None) -> pd.DataFrame:
        data = self.get("/v2/rates/forward", base=base, quote=quote, tenor=_list(tenor) if tenor else None)
        df = pd.DataFrame(data["quotes"])
        for c in df.columns:
            if c.startswith(("spot_", "forward_points_")):
                df[c] = pd.to_numeric(df[c])
        return df
