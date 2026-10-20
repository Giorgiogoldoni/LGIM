#!/usr/bin/env python3
"""Scomposizione del backtest LGIM.

Non cambia le regole del portafoglio: le confronta con varianti e riferimenti sugli STESSI dati, per capire da dove
viene la differenza dai benchmark.

  base               regole attuali
  senza_costi        regole attuali senza costi di transazione (quanto pesa il turnover)
  sempre_investito   stessa classifica e stessi pesi, ma senza filtro di trend, senza stress, senza XEON
                     (il timing aggiunge o toglie valore? la sola selezione per punteggio funziona?)
  universo_pu        tutti gli ETF dell'universo a pesi uguali, ribilanciati ogni mese (la selezione batte il semplice possesso?)
  consigliato        portafoglio statico di L&G (pesi in regole.json), ribilanciato ogni anno

Uso:
  python scomposizione.py                 scarica da Yahoo Finance e scrive data/scomposizione.json e l'Excel
  python scomposizione.py --cache DIR     legge i prezzi da una cartella (test)
"""
import argparse, copy, datetime as dt, json, os, sys

import pandas as pd

QUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, QUI)
import dati  # noqa: E402
import portafoglio as PF  # noqa: E402
import backtest as BT  # noqa: E402

DESCRIZIONI = {
    "base": "Regole attuali del portafoglio",
    "senza_costi": "Regole attuali, senza costi di transazione",
    "sempre_investito": "Stessa classifica e stessi pesi, ma sempre investito: nessun filtro di trend, nessuno stress, nessun XEON",
    "universo_pu": "Tutti gli ETF dell'universo disponibili, a pesi uguali, ribilanciati ogni mese",
    "consigliato": "Portafoglio statico suggerito da L&G (pesi di regole.json), ribilanciato ogni anno",
}
ORDINE = ["base", "senza_costi", "sempre_investito", "universo_pu", "consigliato"]


def variante(R, nome):
    R2 = copy.deepcopy(R); P = R2["parametri"]
    if nome == "senza_costi":
        P["costo_per_lato_pct"] = 0.0
    elif nome == "sempre_investito":
        P["variante"] = {"senza_filtro_trend": True, "senza_stress": True, "senza_paracadute": True}
    return R2


def coppie_date(df, R, inizio, fine):
    """(data di decisione, data di esecuzione) come nel backtest."""
    cal = df[R["parametri"]["backtest"]["calendario"]].index
    dec = BT.date_decisione(cal, inizio or R["parametri"]["backtest"]["inizio"])
    if fine:
        dec = [d for d in dec if d <= pd.Timestamp(fine)]
    out = []
    for d in dec:
        pos = cal.searchsorted(d, side="right")
        if pos < len(cal):
            out.append((d, cal[pos]))
    return out


def simula_statico(df, R, coppie, pesi_a_data, deve_ribilanciare):
    """Portafoglio senza segnali: pesi_a_data(d) -> {ticker: peso 0-1}, deve_ribilanciare(d, precedente) -> bool."""
    costo = R["parametri"]["costo_per_lato_pct"] / 100.0
    unita, ultimo = {}, {}
    valore, prec, serie, turnover = 100.0, None, [], 0.0
    for d, e in coppie:
        prezzi = {t: BT.prezzo_esecuzione(x, d) for t, x in df.items()}
        for t, p in prezzi.items():
            if p is not None:
                ultimo[t] = p
        totale = sum(u * ultimo[t] for t, u in unita.items()) if unita else valore
        pesi = pesi_a_data(d)
        pesi = {t: w for t, w in pesi.items() if prezzi.get(t) is not None}
        s = sum(pesi.values())
        pesi = {t: w / s for t, w in pesi.items()} if s > 0 else {}
        if pesi and (prec is None or set(pesi) != {k for k in prec if k != "_data"} or deve_ribilanciare(d, prec)):
            attuali = {t: u * ultimo[t] for t, u in unita.items()}
            mosso = sum(abs(totale * pesi.get(t, 0.0) - attuali.get(t, 0.0)) for t in set(pesi) | set(attuali))
            totale -= mosso * costo; turnover += mosso / max(totale, 1e-9) * 100
            unita = {t: totale * w / prezzi[t] for t, w in pesi.items()}
            prec = dict(pesi); prec["_data"] = d
        valore = totale
        serie.append((str(e.date()), valore))
    return serie, turnover


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", default=None); ap.add_argument("--inizio", default=None); ap.add_argument("--fine", default=None)
    ap.add_argument("--prova", action="store_true")
    args = ap.parse_args()
    R = PF.carica_regole(); P = R["parametri"]
    universo = [u["ticker"] for u in R["universo"] if u.get("attivo", True)]
    cons = R["consigliato"]["pesi"]
    tickers = sorted(set(universo + [P["paracadute"]] + P["benchmark"] + list(cons)))
    df, falliti = dati.ottieni(tickers, R, args.cache)
    if P["paracadute"] not in df or P["backtest"]["calendario"] not in df:
        sys.exit("ERRORE: mancano i prezzi del paracadute o del calendario di riferimento")
    risultati = {}
    for nome in ("base", "senza_costi", "sempre_investito"):
        print("variante", nome, "...", flush=True)
        risultati[nome] = BT.esegui_backtest(variante(R, nome), df, args.inizio, args.fine)
    base = risultati["base"]
    coppie = coppie_date(df, R, args.inizio, args.fine)
    coppie = [(d, e) for d, e in coppie if str(e.date()) >= base["periodo"]["inizio"]]
    prima = {t: pd.Timestamp(d) for t, d in base["copertura"].items()}
    # universo a pesi uguali: ETF con saldo valido a quella data, ribilanciato al primo giorno di ogni mese
    def pesi_pu(d):
        d = pd.Timestamp(d); return {t: 1.0 for t in universo if t in prima and prima[t] <= d}
    def mensile(d, prec):
        return d.strftime("%Y-%m") != prec["_data"].strftime("%Y-%m")
    serie_pu, turn_pu = simula_statico(df, R, coppie, pesi_pu, mensile)
    # consigliato L&G: pesi fissi tra gli ETF con prezzo disponibile, ribilanciato ogni anno
    def pesi_cons(d):
        d = pd.Timestamp(d); return {t: float(w) for t, w in cons.items() if t in df and df[t].index[0] <= d}
    def annuale(d, prec):
        return d.year != prec["_data"].year
    serie_cs, turn_cs = simula_statico(df, R, coppie, pesi_cons, annuale)
    # serie allineate sulle date di esecuzione del backtest base
    per_data = {p["data"]: {"data": p["data"]} for p in base["serie"]}
    for nome in ("base", "senza_costi", "sempre_investito"):
        for p in risultati[nome]["serie"]:
            if p["data"] in per_data:
                per_data[p["data"]][nome] = p["modello"]
    for nome, serie in (("universo_pu", serie_pu), ("consigliato", serie_cs)):
        if serie:
            b0 = dict(serie).get(base["serie"][0]["data"], serie[0][1])
            for dd, v in serie:
                if dd in per_data:
                    per_data[dd][nome] = round(v / b0 * 100, 2)
    for p in base["serie"]:
        for b in P["benchmark"]:
            if b in p:
                per_data[p["data"]][b] = p[b]
    serie = [per_data[k] for k in sorted(per_data)]
    chiavi = ORDINE + P["benchmark"]
    stats = {}
    for k in chiavi:
        pts = [(p["data"], p[k]) for p in serie if p.get(k) is not None]
        if len(pts) > 2:
            stats[k] = BT.statistiche([v for _, v in pts], [pd.Timestamp(d) for d, _ in pts])
    anni = max((pd.Timestamp(serie[-1]["data"]) - pd.Timestamp(serie[0]["data"])).days / 365.25, 1e-9)
    for nome in ("base", "senza_costi", "sempre_investito"):
        stats[nome]["turnover_annuo"] = risultati[nome]["statistiche"]["modello"]["turnover_annuo"]
        stats[nome]["quota_media_paracadute"] = risultati[nome]["statistiche"]["modello"]["quota_media_paracadute"]
    stats["universo_pu"]["turnover_annuo"] = round(turn_pu / 2 / anni, 1)
    stats["consigliato"]["turnover_annuo"] = round(turn_cs / 2 / anni, 1)
    out = {"generato": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "regole_versione": R["versione"],
           "periodo": base["periodo"], "costo_per_lato_pct": P["costo_per_lato_pct"], "benchmark": P["benchmark"],
           "descrizioni": DESCRIZIONI, "ordine": ORDINE, "serie": serie, "statistiche": stats,
           "anni": BT.rendimenti_per_anno(serie, [k for k in chiavi if k in stats]),
           "consigliato_pesi": cons, "consigliato_valido_dal": R["consigliato"]["valido_dal"],
           "avvisi": (["Yahoo Finance non ha consegnato: " + ", ".join(sorted(falliti))] if falliti else [])}
    print(f"Scomposizione {out['periodo']['inizio']} - {out['periodo']['fine']}")
    for k in chiavi:
        if k in stats:
            print(f"  {k:18s}", stats[k])
    if not args.prova:
        dest = os.environ.get("LGIM_DATA", os.path.join(PF.BASE, "data")); os.makedirs(dest, exist_ok=True)
        with open(os.path.join(dest, "scomposizione.json"), "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, allow_nan=False)
        import esporta
        esporta.scrivi_xlsx_scomposizione(out, os.path.join(dest, "scomposizione_lgim.xlsx"))


if __name__ == "__main__":
    main()
