"""Prezzi giornalieri per il portafoglio LGIM: Yahoo Finance in produzione, cartella di cache nei test."""
import json
import os

import pandas as pd


def pulisci(d, salto_max, salto_min):
    """Toglie righe senza prezzo valido, duplicati e barre con salti anomali rispetto al giorno prima."""
    d = d[~d.index.duplicated(keep="last")].sort_index()
    d = d.dropna(subset=["Close"])
    d = d[(d.Close > 0) & (d.High > 0) & (d.Low > 0) & (d.Open > 0)]
    r = d.Close / d.Close.shift(1)
    return d[~((r > salto_max) | (r < salto_min))]


def da_yahoo(ticker, anni, salto_max, salto_min):
    import yfinance as yf
    d = yf.Ticker(ticker).history(period=f"{anni}y", interval="1d", auto_adjust=False)
    if d is None or d.empty:
        return None
    d.index = pd.to_datetime(d.index).tz_localize(None).normalize()
    return pulisci(d[["Open", "High", "Low", "Close", "Volume"]], salto_max, salto_min)


def da_cache(ticker, cartella, salto_max, salto_min):
    p = os.path.join(cartella, ticker.replace(".", "_") + ".json")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        j = json.load(f)
    d = pd.DataFrame({"Open": j["o"], "High": j["h"], "Low": j["l"], "Close": j["c"], "Volume": j.get("v") or [0] * len(j["c"])},
                     index=pd.to_datetime(j["dates"]))
    return pulisci(d, salto_max, salto_min)


def ottieni(tickers, R, cartella_cache=None, pausa=0.5):
    """Ritorna ({ticker: DataFrame}, {ticker: motivo del fallimento})."""
    import time
    c = R["parametri"]["download"]
    out, falliti = {}, {}
    for t in tickers:
        try:
            d = da_cache(t, cartella_cache, c["salto_max"], c["salto_min"]) if cartella_cache else \
                da_yahoo(t, c["anni_storia"], c["salto_max"], c["salto_min"])
        except Exception as e:  # rete, limiti di Yahoo, ticker sconosciuto
            falliti[t] = repr(e)[:120]; continue
        if d is None or len(d) < c["min_barre"]:
            falliti[t] = "nessun dato" if d is None else f"storia insufficiente ({len(d)} barre)"
        else:
            out[t] = d
        if not cartella_cache:
            time.sleep(pausa)
    return out, falliti
