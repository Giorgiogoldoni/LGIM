"""Indicatori e saldi del portafoglio LGIM.

Le formule sono le stesse di investing (compute_tech.py), ma qui le soglie vengono da regole.json:
cambiare investing non cambia questo portafoglio. Il saldo usa un set FISSO di voti: se ne manca anche
uno (storia insufficiente) l'orizzonte non e' disponibile.
"""
import numpy as np
import pandas as pd

AGG = {"Open": "first", "High": "max", "Low": "min", "Close": "last"}
MEDIE = (5, 10, 20, 50)


def wil(s, n):
    return s.ewm(alpha=1 / n, adjust=False, min_periods=n).mean()


def nn(v):
    return v is None or pd.isna(v)


def serie_indicatori(d):
    """Indicatori per ogni barra (ognuno dipende solo dalle barre precedenti)."""
    h, l, c = d.High, d.Low, d.Close
    pc = c.shift(1); up = c.diff()
    rsi = 100 - 100 / (1 + wil(up.clip(lower=0), 14) / wil((-up).clip(lower=0), 14))
    k = (100 * (c - l.rolling(9).min()) / (h.rolling(9).max() - l.rolling(9).min())).rolling(6).mean()
    srsi = 100 * (rsi - rsi.rolling(14).min()) / (rsi.rolling(14).max() - rsi.rolling(14).min())
    mac = c.ewm(span=12, adjust=False).mean() - c.ewm(span=26, adjust=False).mean()
    diff = mac - mac.ewm(span=9, adjust=False).mean()
    tr = pd.concat([h - l, (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1); atr = wil(tr, 14)
    pdm = h.diff().where((h.diff() > -l.diff()) & (h.diff() > 0), 0.0)
    mdm = (-l.diff()).where((-l.diff() > h.diff()) & (-l.diff() > 0), 0.0)
    pdi = 100 * wil(pdm, 14) / atr; mdi = 100 * wil(mdm, 14) / atr
    adx = wil(100 * (pdi - mdi).abs() / (pdi + mdi), 14)
    tp = (h + l + c) / 3
    cci = (tp - tp.rolling(14).mean()) / (0.015 * tp.rolling(14).apply(lambda x: np.abs(x - x.mean()).mean(), raw=True))
    roc = 100 * (c / c.shift(12) - 1)
    lo_ = pd.concat([l, pc], axis=1).min(axis=1); bp = c - lo_
    t2 = pd.concat([h, pc], axis=1).max(axis=1) - lo_
    av = lambda n: bp.rolling(n).sum() / t2.rolling(n).sum()
    uo = 100 * (4 * av(7) + 2 * av(14) + av(28)) / 7
    e13 = c.ewm(span=13, adjust=False).mean()
    S = pd.DataFrame({"close": c, "rsi": rsi, "k": k, "srsi": srsi, "macd_diff": diff, "adx": adx, "pdi": pdi, "mdi": mdi,
                      "cci": cci, "roc": roc, "uo": uo, "e13": e13, "bull": h - e13, "bear": l - e13}, index=d.index)
    for n in MEDIE:
        S[f"sma{n}"] = c.rolling(n).mean()
        S[f"ema{n}"] = c.ewm(span=n, adjust=False, min_periods=n).mean()
    return S


def _vote(v, lo, hi):
    return None if nn(v) else "Compra" if v >= hi else "Vendi" if v <= lo else "Neutrale"


def _sgn(v):
    return None if nn(v) else "Compra" if v > 0 else "Vendi" if v < 0 else "Neutrale"


def voti_riga(S, i, soglie):
    """Voti alla barra i (posizione). Ritorna ({nome: voto}, {periodo: (v_sma, v_ema)})."""
    r = S.iloc[i]; p = S.iloc[i - 1] if i > 0 else None
    adx_v = None if nn(r.adx) else ("Neutrale" if r.adx <= soglie["adx"] else "Compra" if r.pdi > r.mdi else "Vendi")
    el = "Neutrale"
    if p is not None:
        if r.e13 > p.e13 and r.bear < 0 and r.bear > p.bear:
            el = "Compra"
        elif r.e13 < p.e13 and r.bull > 0 and r.bull < p.bull:
            el = "Vendi"
    ind = {"RSI(14)": _vote(r.rsi, *soglie["rsi"]), "STOCH(9,6)": _vote(r.k, *soglie["stoch"]),
           "STOCHRSI(14)": _vote(r.srsi, *soglie["stoch"]), "MACD(12,26,9)": _sgn(r.macd_diff), "ADX(14)": adx_v,
           "CCI(14)": _sgn(r.cci), "ROC(12)": _sgn(r.roc), "Ultimate Oscillator": None if nn(r.uo) else _sgn(r.uo - 50),
           "Bull/Bear Power(13)": el}
    med = {}
    for n in MEDIE:
        med[n] = (None if nn(r[f"sma{n}"]) else "Compra" if r.close > r[f"sma{n}"] else "Vendi",
                  None if nn(r[f"ema{n}"]) else "Compra" if r.close > r[f"ema{n}"] else "Vendi")
    return ind, med


def saldo_fisso(ind, med, indicatori, periodi):
    voti = [ind.get(n) for n in indicatori]
    for p in periodi:
        voti += list(med.get(p, (None, None)))
    if any(v is None for v in voti):
        return None
    return round((voti.count("Compra") - voti.count("Vendi")) / len(voti) * 100, 1)


def saldi_a_data(df, R, data=None, S_D=None):
    """Saldi giornaliero, settimanale e mensile usando solo le barre fino a `data` (incluso).

    df: Open/High/Low/Close con indice data. S_D: indicatori giornalieri gia' calcolati su tutto df (facoltativo,
    accelera il backtest: ogni valore dipende solo dal passato, quindi e' identico).
    """
    P = R["parametri"]; V = P["voti"]; minb = P["download"]["min_barre"]
    d = df if data is None else df[df.index <= pd.Timestamp(data)]
    if len(d) < minb:
        return None
    out = {"D": None, "W": None, "M": None, "ultima": str(d.index[-1].date()), "prezzo": float(d.Close.iloc[-1])}
    S = S_D.loc[:d.index[-1]] if S_D is not None else serie_indicatori(d)
    ind, med = voti_riga(S, len(S) - 1, P["soglie"])
    out["D"] = saldo_fisso(ind, med, V["indicatori"], V["medie_giornaliero_settimanale"])
    for chiave, regola, periodi in (("W", "W-FRI", V["medie_giornaliero_settimanale"]), ("M", "ME", V["medie_mensile"])):
        f = d.resample(regola).agg(AGG).dropna()
        if len(f) >= minb:
            Sf = serie_indicatori(f)
            ind, med = voti_riga(Sf, len(Sf) - 1, P["soglie"])
            out[chiave] = saldo_fisso(ind, med, V["indicatori"], periodi)
    return out
