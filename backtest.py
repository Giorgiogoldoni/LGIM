#!/usr/bin/env python3
"""Backtest del portafoglio LGIM.

Riesegue le stesse regole del portafoglio reale (funzione passo() di portafoglio.py) sul passato:
- ogni decisione e' presa sull'ultima seduta della settimana usando SOLO i prezzi fino a quel giorno;
- si esegue alla chiusura della seduta successiva, con il costo per lato di regole.json;
- un ETF e' disponibile solo da quando ha storia sufficiente per il set fisso di voti.

Uso:
  python backtest.py                  scarica da Yahoo Finance e scrive data/backtest.json e l'Excel
  python backtest.py --cache DIR      legge i prezzi da una cartella (test)
  python backtest.py --inizio 2022-01-07 --fine 2024-06-28 --prova   solo un sotto-periodo (progetto/verifica)
"""
import argparse, datetime as dt, json, math, os, sys

import pandas as pd

QUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, QUI)
import dati, indicatori  # noqa: E402
import portafoglio as PF  # noqa: E402


def date_decisione(cal, inizio):
    s = pd.Series(cal, index=cal)
    ultimi = s.groupby(cal.to_period("W-FRI")).last()
    return [pd.Timestamp(d) for d in ultimi.values if pd.Timestamp(d) >= pd.Timestamp(inizio)]


def prezzo_esecuzione(d, data):
    """Chiusura della prima seduta successiva a `data` (al massimo 7 giorni dopo), altrimenti None."""
    i = d.index.searchsorted(data, side="right")
    if i < len(d) and (d.index[i] - data).days <= 7:
        return float(d.Close.iloc[i])
    return None


def statistiche(valori, date):
    s = pd.Series(valori, dtype=float)
    anni = max((date[-1] - date[0]).days / 365.25, 1e-9)
    rend = s.pct_change().dropna()
    return {"rendimento_totale": round(float((s.iloc[-1] / s.iloc[0] - 1) * 100), 1),
            "rendimento_annuo": round(float(((s.iloc[-1] / s.iloc[0]) ** (1 / anni) - 1) * 100), 1),
            "volatilita_annua": round(float(rend.std()) * math.sqrt(52) * 100, 1) if len(rend) > 1 else None,
            "drawdown_massimo": round(float((s / s.cummax() - 1).min()) * 100, 1)}


def rendimenti_per_anno(serie, chiavi):
    anni = sorted({p["data"][:4] for p in serie})
    out = []
    for a in anni:
        riga = {"anno": a, "parziale": a in (anni[0], anni[-1])}
        prima = [p for p in serie if p["data"][:4] < a]
        base = prima[-1] if prima else serie[0]
        fine = [p for p in serie if p["data"][:4] == a][-1]
        for k in chiavi:
            if base.get(k) and fine.get(k):
                riga[k] = round((fine[k] / base[k] - 1) * 100, 1)
        out.append(riga)
    return out


def esegui_backtest(R, df, inizio=None, fine=None):
    P = R["parametri"]; pc_t = P["paracadute"]
    universo = [u["ticker"] for u in R["universo"] if u.get("attivo", True)]
    cal = df[P["backtest"]["calendario"]].index
    S_D = {t: indicatori.serie_indicatori(df[t]) for t in universo if t in df}
    decisioni = date_decisione(cal, inizio or P["backtest"]["inizio"])
    if fine:
        decisioni = [d for d in decisioni if d <= pd.Timestamp(fine)]
    if len(decisioni) < 2:
        sys.exit("ERRORE: poche date di decisione: controlla la data di inizio e i dati")
    stato = PF.stato_nuovo(decisioni[0].date())
    chiavi = ["modello"] + P["benchmark"]
    serie, storico, copertura = [], [], {}
    somma_turnover = somma_xeon = somma_sat = n_pass = 0.0
    contatori = {}
    base_bm = {}
    for d in decisioni:
        pos = cal.searchsorted(d, side="right")
        if pos >= len(cal):
            break
        e = cal[pos]
        val = {t: (indicatori.saldi_a_data(df[t], R, d, S_D[t]) if t in df else None) for t in universo}
        for t, v in val.items():
            if v and v["D"] is not None and v["W"] is not None:
                copertura.setdefault(t, str(d.date()))
        cands, rif = PF.costruisci_candidati(R, val)
        prezzi = {t: prezzo_esecuzione(d_, d) for t, d_ in df.items()}
        prezzi = {t: p for t, p in prezzi.items() if p is not None}
        if pc_t not in prezzi:
            continue
        r = PF.passo(R, cands, prezzi, stato, e.date())
        contatori[r["azione"]] = contatori.get(r["azione"], 0) + 1
        somma_turnover += r["turnover_pct"]; n_pass += 1
        somma_xeon += r["pesi"].get(pc_t, 0.0)
        per_t = {c["ticker"]: c for c in cands}
        somma_sat += sum(w for t, w in r["pesi"].items() if per_t.get(t) and per_t[t]["parte"] == "satellite")
        punto = {"data": str(e.date()), "dati_al": rif, "azione": r["azione"], "modello": round(r["valore"], 2)}
        for b in P["benchmark"]:
            if b in prezzi:
                base_bm.setdefault(b, prezzi[b]); punto[b] = round(prezzi[b] / base_bm[b] * 100, 2)
        serie.append(punto)
        if r["azione"] != "MANTENIMENTO":
            storico.append({"data": str(e.date()), "dati_al": rif, "azione": r["azione"], "motivo": r["motivo"],
                            "movimenti": r["movimenti"]})
    base = serie[0]["modello"]
    for p in serie:
        p["modello"] = round(p["modello"] / base * 100, 2)
    date = [pd.Timestamp(p["data"]) for p in serie]
    stats = {k: statistiche([p[k] for p in serie if k in p], [d for d, p in zip(date, serie) if k in p]) for k in chiavi}
    anni = max((date[-1] - date[0]).days / 365.25, 1e-9)
    stats["modello"].update({"turnover_annuo": round(somma_turnover / 2 / anni, 1), "quota_media_paracadute": round(somma_xeon / n_pass, 1),
                             "quota_media_satellite": round(somma_sat / n_pass, 1), "decisioni": int(n_pass)})
    return {"generato": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "regole_versione": R["versione"],
            "periodo": {"inizio": serie[0]["data"], "fine": serie[-1]["data"]}, "costo_per_lato_pct": P["costo_per_lato_pct"],
            "benchmark": P["benchmark"], "serie": serie, "statistiche": stats, "anni": rendimenti_per_anno(serie, chiavi),
            "azioni": contatori, "storico": storico[::-1], "copertura": copertura}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", default=None); ap.add_argument("--inizio", default=None); ap.add_argument("--fine", default=None); ap.add_argument("--prova", action="store_true")
    args = ap.parse_args()
    R = PF.carica_regole(); P = R["parametri"]
    universo = [u["ticker"] for u in R["universo"] if u.get("attivo", True)]
    tickers = universo + [P["paracadute"]] + P["benchmark"]
    df, falliti = dati.ottieni(tickers, R, args.cache)
    if P["paracadute"] not in df or P["backtest"]["calendario"] not in df:
        sys.exit("ERRORE: mancano i prezzi del paracadute o del calendario di riferimento")
    bt = esegui_backtest(R, df, args.inizio, args.fine)
    bt["avvisi"] = (["Yahoo Finance non ha consegnato: " + ", ".join(sorted(falliti))] if falliti else []) + \
                   ["ETF senza storia sufficiente nel periodo: " + ", ".join(t for t in universo if t not in bt["copertura"])] \
                   * bool([t for t in universo if t not in bt["copertura"]])
    s = bt["statistiche"]
    print(f"Backtest {bt['periodo']['inizio']} - {bt['periodo']['fine']}, costo {bt['costo_per_lato_pct']}% per lato, {s['modello']['decisioni']} decisioni")
    for k, v in s.items():
        print(f"  {k:10s}", v)
    if not args.prova and (args.fine or args.inizio):
        print("Periodo personalizzato: risultati non salvati (usa --prova o lancia senza --inizio/--fine per il risultato ufficiale)")
    elif not args.prova:
        dest = os.environ.get("LGIM_DATA", os.path.join(PF.BASE, "data")); os.makedirs(dest, exist_ok=True)
        with open(os.path.join(dest, "backtest.json"), "w", encoding="utf-8") as f:
            json.dump(bt, f, ensure_ascii=False, allow_nan=False)
        import esporta
        esporta.scrivi_xlsx_backtest(bt, R, os.path.join(dest, "backtest_lgim.xlsx"))


if __name__ == "__main__":
    main()
