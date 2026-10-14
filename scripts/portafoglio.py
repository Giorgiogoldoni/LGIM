#!/usr/bin/env python3
"""Portafoglio modello LGIM (autonomo).

Scarica i prezzi da Yahoo Finance, calcola i saldi con un set fisso di voti, applica regole.json e aggiorna
data/stato.json, data/portafoglio.json, l'Excel e il PDF delle regole.

Uso:
  python scripts/portafoglio.py                   esecuzione normale
  python scripts/portafoglio.py --forza-mensile   forza il ribilanciamento mensile
  python scripts/portafoglio.py --prova           calcola e stampa, non scrive nulla
  python scripts/portafoglio.py --cache DIR       legge i prezzi da una cartella invece che da Yahoo (test)
  python scripts/portafoglio.py --oggi 2026-11-02 data di esecuzione (solo per test)
"""
import argparse, collections, datetime as dt, json, os, sys

QUI = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(QUI)
DATA = os.environ.get("LGIM_DATA", os.path.join(BASE, "data"))
sys.path.insert(0, QUI)
import dati, indicatori  # noqa: E402


def carica_regole():
    with open(os.path.join(BASE, "regole.json"), encoding="utf-8") as f:
        return json.load(f)


def leggi_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def media(valori):
    v = [x for x in valori if x is not None]
    return round(sum(v) / len(v), 1) if v else None


# ------------------------------------------------------------------ candidati
def costruisci_candidati(R, valutazioni):
    """valutazioni: {ticker: saldi_a_data(...) oppure None se mancano i dati}."""
    out = []
    for u in R["universo"]:
        c = dict(u); t = u["ticker"]
        c["gruppo"] = u.get("area") or u.get("famiglia")
        c.update(stato="ok", motivo=None, prezzo=None, D=None, W=None, M=None, ultima=None, saldo_medio=None,
                 punteggio=None, ok_ranking=False, ok_ingresso=False, vecchio=False)
        v = valutazioni.get(t)
        if not u.get("attivo", True):
            c["stato"], c["motivo"] = "escluso", "escluso dalle regole"
        elif v is None:
            c["stato"], c["motivo"] = "senza_dati", "dati non disponibili da Yahoo Finance (o storia insufficiente)"
        else:
            c.update(D=v["D"], W=v["W"], M=v["M"], ultima=v["ultima"], prezzo=v["prezzo"])
            if c["D"] is None or c["W"] is None:
                c["stato"], c["motivo"] = "storia_insufficiente", "storia insufficiente per il set fisso di voti (giornaliero o settimanale)"
            else:
                c["saldo_medio"] = media([c["D"], c["W"], c["M"]])
                c["punteggio"] = media([c["W"], c["M"]])
                c["ok_ranking"] = c["saldo_medio"] > 0
                if not c["ok_ranking"]:
                    c["motivo"] = "saldo medio non positivo"
        out.append(c)
    date = collections.Counter(c["ultima"] for c in out if c["ultima"])
    rif = max(date.items(), key=lambda kv: (kv[1], kv[0]))[0] if date else None
    for c in out:
        c["vecchio"] = bool(rif and c["ultima"] and c["ultima"] < rif)
        c["ok_ingresso"] = bool(c["ok_ranking"] and (c["D"] or 0) > 0 and not c["vecchio"])
        if c["vecchio"] and c["motivo"] is None:
            c["motivo"] = f"ultima barra {c['ultima']} più vecchia del riferimento {rif}"
    return out, rif


# ------------------------------------------------------------------- selezione
def entra_ok(c, scelti, regole):
    """True se aggiungere c a `scelti` rispetta i limiti di gruppo; altrimenti (False, motivo)."""
    esplicite = [r for r in regole if r.get("valori")]
    for r in esplicite:
        if c[r["campo"]] in r["valori"] and sum(1 for s in scelti if s[r["campo"]] in r["valori"]) >= r["max"]:
            return False, "limite di %d su %s" % (r["max"], " e ".join(r["valori"]))
    for r in (x for x in regole if not x.get("valori")):
        coperto = any(c[r["campo"]] in e["valori"] for e in esplicite if e["campo"] == r["campo"])
        if not coperto and sum(1 for s in scelti if s[r["campo"]] == c[r["campo"]]) >= r["max"]:
            return False, "limite di %d per %s %s" % (r["max"], r["campo"], c[r["campo"]])
    return True, None


def scegli_parte(parte, cands, tenuti, R, oggi, ingressi, congelati):
    P = R["parametri"]; pc = P["posizioni"][parte]; lim = P["limiti"][parte]
    per_t = {c["ticker"]: c for c in cands}; held = set(tenuti)
    # chi non puo' entrare non occupa posti nei limiti di gruppo
    elig = [c for c in cands if c["parte"] == parte and c.get("attivo", True) and
            ((c["ticker"] in held and c["ok_ranking"]) or (c["ticker"] not in held and c["ok_ingresso"]))]
    elig.sort(key=lambda c: (-c["punteggio"], -c["saldo_medio"], c["ticker"]))
    lista, meta = [], {}
    for c in elig:
        ok, mot = entra_ok(c, lista, lim)
        if not ok:
            meta[c["ticker"]] = {"scartato": mot}; continue
        lista.append(c)
    rank = {c["ticker"]: i + 1 for i, c in enumerate(lista)}
    pos = []
    for t in tenuti:
        c = per_t.get(t)
        if c is None or c["parte"] != parte:
            continue
        if t in congelati:
            pos.append(t)
            meta[t] = {"nota": "dati non aggiornati: resta fermo" if c["vecchio"] and c["punteggio"] is not None
                       else "dati non disponibili: peso mantenuto"}
            continue
        if t in rank and rank[t] <= pc["esce_oltre"]:
            pos.append(t); continue
        giorni = (oggi - dt.date.fromisoformat(ingressi.get(t, oggi.isoformat()))).days
        if c["ok_ranking"] and giorni < P["permanenza_minima_settimane"] * 7:
            pos.append(t); meta[t] = {"nota": f"permanenza minima ({giorni} giorni)"}; continue
        if not c["ok_ranking"]:
            uscita = "saldo medio non positivo"
        elif t in rank:
            uscita = f"posizione {rank[t]}, oltre la {pc['esce_oltre']}"
        else:
            uscita = meta.get(t, {}).get("scartato", "limite di gruppo")
        meta[t] = {"uscita": uscita}
    for c in lista:
        if len(pos) >= pc["max"]:
            break
        if c["ticker"] in pos or not c["ok_ingresso"] or rank[c["ticker"]] > pc["entra_entro"]:
            continue
        if entra_ok(c, [per_t[x] for x in pos], lim)[0]:
            pos.append(c["ticker"]); meta[c["ticker"]] = {"ingresso": f"posizione {rank[c['ticker']]} nella parte {parte}"}
    for c in lista:
        if len(pos) >= pc["min"]:
            break
        if c["ticker"] in pos or not c["ok_ingresso"]:
            continue
        if P.get("riempimento_richiede_settimanale_positivo") and (c["W"] or 0) <= 0:
            continue
        if entra_ok(c, [per_t[x] for x in pos], lim)[0]:
            pos.append(c["ticker"])
            meta[c["ticker"]] = {"ingresso": f"posizione {rank[c['ticker']]}, per raggiungere il minimo di {pc['min']}"}
    return pos, meta


def pesi_parte(parte, pos, cands, P, congelati, cw):
    per_t = {c["ticker"]: c for c in cands}; pc = P["posizioni"][parte]
    budget = P["core_target"] if parte == "core" else P["satellite"]["target"]
    fissi = {t: cw.get(t, 0.0) for t in pos if t in congelati}
    liberi = [t for t in pos if t not in congelati]
    B = max(0.0, budget - sum(fissi.values())) * min(1.0, len(pos) / pc["min"])
    p = {t: max(per_t[t]["punteggio"] if per_t[t]["punteggio"] is not None else 1.0, 1.0) for t in liberi}
    w = {t: 0.0 for t in liberi}
    if liberi and B > 0:
        tot = sum(p.values()); w = {t: B * p[t] / tot for t in liberi}
        for _ in range(10):
            sopra = [t for t in liberi if w[t] > P["tetto_peso"] + 1e-9]
            if not sopra:
                break
            eccesso = sum(w[t] - P["tetto_peso"] for t in sopra)
            for t in sopra:
                w[t] = float(P["tetto_peso"])
            sotto = [t for t in liberi if w[t] < P["tetto_peso"] - 1e-9]
            if not sotto:
                break
            ps = sum(p[t] for t in sotto)
            for t in sotto:
                w[t] += eccesso * p[t] / ps
    w.update(fissi)
    return w


def obiettivo(cands, tenuti, R, oggi, ingressi, cw):
    P = R["parametri"]; per_t = {c["ticker"]: c for c in cands}
    congelati = {t for t in tenuti if t in per_t and (per_t[t]["punteggio"] is None or per_t[t]["vecchio"])}
    tgt, meta = {}, {}
    for parte in ("core", "satellite"):
        pos, m = scegli_parte(parte, cands, tenuti, R, oggi, ingressi, congelati)
        meta.update(m); tgt.update(pesi_parte(parte, pos, cands, P, congelati, cw))
    tgt = {t: round(v, 2) for t, v in tgt.items() if v > 0.004}
    tgt[P["paracadute"]] = round(max(0.0, 100.0 - sum(tgt.values())), 2)
    return tgt, meta


# ----------------------------------------------------------------------- stato
def stato_nuovo(oggi):
    return {"versione": 2, "creato": oggi.isoformat(), "valore": 100.0, "posizioni": {}, "ultimo_mensile": None,
            "benchmark_base": {}, "nav_serie": [], "storico": []}


def passo(R, cands, prezzi, stato, oggi, forza_mensile=False):
    """Un'esecuzione del portafoglio: decide, muove, aggiorna lo stato. Usato da produzione e backtest."""
    P = R["parametri"]; pc_t = P["paracadute"]; per_t = {c["ticker"]: c for c in cands}
    avvisi = []
    nuovo = stato.get("ultimo_mensile") is None and not stato["posizioni"]
    val, pr = {}, {}
    for t, pz in stato["posizioni"].items():
        p = prezzi.get(t)
        if p is None:
            avvisi.append(f"Prezzo non disponibile per {t}: usato l'ultimo noto"); p = pz["prezzo"]
        pr[t] = p; val[t] = pz["unita"] * p
    totale = sum(val.values()) if val else stato["valore"]
    cw = {t: v / totale * 100 for t, v in val.items()} if val else {}
    tenuti = [t for t in cw if t != pc_t and cw[t] > 0]
    ingressi = {t: pz.get("ingresso", oggi.isoformat()) for t, pz in stato["posizioni"].items()}
    eq = sum(cw[t] for t in tenuti)
    negativi = [t for t in tenuti if per_t.get(t) and per_t[t]["D"] is not None and per_t[t]["W"] is not None
                and per_t[t]["D"] < 0 and per_t[t]["W"] < 0 and not per_t[t]["vecchio"]]
    quota_stress = (sum(cw[t] for t in negativi) / eq * 100) if eq > 0 else 0.0
    stress = eq > 0 and quota_stress >= P["stress_quota_peso"]
    tgt, meta = obiettivo(cands, tenuti, R, oggi, ingressi, cw)
    mese = oggi.strftime("%Y-%m")
    if nuovo:
        azione, motivo, nw = "INIZIO", "prima esecuzione: il portafoglio adotta il target", dict(tgt)
    elif stress:
        azione = "STRESS"
        motivo = f"{quota_stress:.0f}% del peso azionario con saldo giornaliero e settimanale negativi"
        nw = {t: (0.0 if t in negativi else cw[t]) for t in tenuti}
        nw[pc_t] = 100.0 - sum(nw.values())
        for t in negativi:
            meta.setdefault(t, {})["uscita"] = "stress: saldo giornaliero e settimanale negativi"
    elif forza_mensile or stato.get("ultimo_mensile") != mese:
        azione, motivo, nw = "MENSILE", "ribilanciamento mensile al target", dict(tgt)
    else:
        sat_w = sum(cw.get(t, 0) for t in tenuti if per_t.get(t) and per_t[t]["parte"] == "satellite")
        fuori = eq > 0 and sat_w > 0 and not (P["satellite"]["min"] <= sat_w <= P["satellite"]["max"])
        dev = {t: abs(tgt.get(t, 0.0) - cw.get(t, 0.0)) for t in set(tgt) | set(tenuti) if t != pc_t}
        sopra = [t for t, d in dev.items() if d >= P["soglia_ribilanciamento"]]
        if fuori:
            azione, motivo, nw = "SETTIMANALE", f"satellite al {sat_w:.1f}%, fuori dalla banda {P['satellite']['min']}-{P['satellite']['max']}%", dict(tgt)
        elif sopra:
            azione, motivo = "SETTIMANALE", "scostamento di almeno %g punti su %s" % (P["soglia_ribilanciamento"], ", ".join(sorted(sopra)))
            nw = {t: (tgt.get(t, 0.0) if t in sopra else cw.get(t, 0.0)) for t in set(tgt) | set(tenuti) if t != pc_t}
            eccesso = sum(nw.values()) - 100.0
            if eccesso > 0:  # XEON non basta: si finanzia riducendo, pro rata, le posizioni sopra target non toccate
                sopra_tgt = {t: cw.get(t, 0.0) - tgt.get(t, 0.0) for t in nw
                             if t not in sopra and cw.get(t, 0.0) > tgt.get(t, 0.0)}
                tot_s = sum(sopra_tgt.values())
                if tot_s > 0:
                    prendi = min(eccesso, tot_s)
                    for t, x in sopra_tgt.items():
                        nw[t] -= prendi * x / tot_s
                    eccesso -= prendi
            if eccesso > 1e-9:  # ancora insufficiente: si ridimensionano gli aumenti
                aumenti = {t: nw[t] - cw.get(t, 0.0) for t in nw if nw[t] > cw.get(t, 0.0)}
                tot_a = sum(aumenti.values())
                for t, a in aumenti.items():
                    nw[t] -= eccesso * a / tot_a
            nw[pc_t] = max(0.0, 100.0 - sum(nw.values()))
        else:
            azione, motivo, nw = "MANTENIMENTO", "nessuno scostamento rilevante", dict(cw)
    nw = {t: round(v, 4) for t, v in nw.items() if v > 0.004}
    if azione == "SETTIMANALE" and all(abs(nw.get(t, 0.0) - cw.get(t, 0.0)) < 0.05 for t in set(nw) | set(cw)):
        azione, motivo, nw = "MANTENIMENTO", "scostamento non finanziabile: nessun movimento", dict(cw)
    movimenti, turnover = [], 0.0
    if azione != "MANTENIMENTO":
        costo = P["costo_per_lato_pct"] / 100.0
        turnover = sum(abs(totale * nw.get(t, 0) / 100 - val.get(t, 0.0)) for t in set(nw) | set(val))
        dopo = totale - turnover * costo
        for t in sorted(set(nw) | set(cw)):
            a, b = cw.get(t, 0.0), nw.get(t, 0.0)
            if abs(b - a) < 0.05:
                continue
            tipo = "ENTRA" if a == 0 else "ESCE" if b == 0 else "AUMENTA" if b > a else "RIDUCI"
            m = meta.get(t, {})
            mot = m.get("ingresso") or m.get("uscita") or (azione.lower() if t != pc_t else
                                                         "quota non assegnata" if tipo in ("ENTRA", "AUMENTA") else "riallocazione")
            movimenti.append({"ticker": t, "nome": (per_t.get(t) or {}).get("nome") or t, "azione": tipo,
                              "da": round(a, 1), "a": round(b, 1), "motivo": mot})
        posizioni = {}
        for t, w in nw.items():
            p = pr.get(t) or prezzi.get(t)
            if p is None:
                avvisi.append(f"Prezzo non disponibile per {t}: posizione non aperta"); continue
            vecchio = stato["posizioni"].get(t)
            posizioni[t] = {"unita": dopo * w / 100 / p, "prezzo": p, "ingresso": vecchio["ingresso"] if vecchio else oggi.isoformat()}
        stato["posizioni"] = posizioni; stato["valore"] = dopo
        if azione in ("INIZIO", "MENSILE"):
            stato["ultimo_mensile"] = mese
        pesi_fin, valore_fin = nw, dopo
    else:
        for t, pz in stato["posizioni"].items():
            pz["prezzo"] = pr[t]
        stato["valore"] = totale; pesi_fin, valore_fin = cw, totale
    return {"azione": azione, "motivo": motivo, "movimenti": movimenti, "stress": bool(stress), "quota_stress": quota_stress,
            "tgt": tgt, "meta": meta, "pesi": pesi_fin, "valore": valore_fin, "avvisi": avvisi,
            "turnover_pct": turnover / totale * 100 if totale else 0.0}


# ------------------------------------------------------------------- produzione
def esegui_live(args):
    R = carica_regole(); P = R["parametri"]
    oggi = dt.date.fromisoformat(args.oggi) if args.oggi else dt.datetime.now(dt.timezone.utc).date()
    universo = [u["ticker"] for u in R["universo"] if u.get("attivo", True)]
    tickers = universo + [P["paracadute"]] + P["benchmark"]
    df, falliti = dati.ottieni(tickers, R, args.cache)
    quota = len([t for t in tickers if t in df]) / len(tickers) * 100
    if quota < P["download"]["quota_minima_dati"]:
        sys.exit(f"ERRORE: dati disponibili solo per il {quota:.0f}% dei ticker: esecuzione interrotta senza modifiche")
    if P["paracadute"] not in df:
        sys.exit(f"ERRORE: nessun prezzo per il paracadute {P['paracadute']}: esecuzione interrotta senza modifiche")
    val = {t: (indicatori.saldi_a_data(df[t], R) if t in df else None) for t in universo}
    cands, rif = costruisci_candidati(R, val)
    per_t = {c["ticker"]: c for c in cands}
    prezzi = {t: float(d.Close.iloc[-1]) for t, d in df.items()}
    stato = leggi_json(os.path.join(DATA, "stato.json"), None) or stato_nuovo(oggi)
    r = passo(R, cands, prezzi, stato, oggi, args.forza_mensile)
    avvisi = list(r["avvisi"])
    bm = {}
    for b in P["benchmark"]:
        if b not in prezzi:
            avvisi.append(f"Prezzo del benchmark {b} non disponibile"); continue
        stato["benchmark_base"].setdefault(b, prezzi[b]); bm[b] = round(prezzi[b] / stato["benchmark_base"][b] * 100, 2)
    punto = {"data": oggi.isoformat(), "dati_al": rif, "nav": round(r["valore"], 2), **bm, "azione": r["azione"]}
    stato["nav_serie"] = ([x for x in stato["nav_serie"] if x["data"] != punto["data"]] + [punto])[-400:]
    if r["azione"] != "MANTENIMENTO":
        stato["storico"] = ([{"data": oggi.isoformat(), "dati_al": rif, "azione": r["azione"], "motivo": r["motivo"],
                              "movimenti": r["movimenti"]}] + stato["storico"])[:200]
    stato["ultima_esecuzione"] = oggi.isoformat()
    if falliti:
        avvisi.append("Yahoo Finance non ha consegnato %d ticker: %s" % (len(falliti), ", ".join(sorted(falliti))))
    vecchi = [c["ticker"] for c in cands if c["vecchio"]]
    if vecchi:
        avvisi.append("ETF con barra più vecchia del riferimento %s: %s" % (rif, ", ".join(vecchi)))
    pesi = r["pesi"]; pc_t = P["paracadute"]
    righe = []
    for t, w in sorted(pesi.items(), key=lambda kv: -kv[1]):
        c = per_t.get(t)
        righe.append({"ticker": t, "nome": (c or {}).get("nome") or t, "isin": (c or {}).get("isin"),
                      "parte": (c or {}).get("parte", "paracadute"), "gruppo": (c or {}).get("gruppo"),
                      "peso": round(w, 1), "target": r["tgt"].get(t, 0.0), "D": (c or {}).get("D"), "W": (c or {}).get("W"),
                      "M": (c or {}).get("M"), "punteggio": (c or {}).get("punteggio"),
                      "ingresso": stato["posizioni"].get(t, {}).get("ingresso"), "prezzo": prezzi.get(t),
                      "dati": "aggiornati" if c and c["ultima"] and not c["vecchio"] else ("non aggiornati" if c and c["vecchio"] else "n.d.")})
    cand_out = []
    for c in cands:
        m = r["meta"].get(c["ticker"], {})
        in_p = c["ticker"] in pesi and c["ticker"] != pc_t
        stato_c = "in portafoglio" if in_p else ("escluso" if c["stato"] != "ok" else
                  "qualificato, non selezionato" if c["ok_ranking"] else "non qualificato")
        cand_out.append({"ticker": c["ticker"], "nome": c["nome"], "isin": c["isin"], "parte": c["parte"], "gruppo": c["gruppo"],
                         "D": c["D"], "W": c["W"], "M": c["M"], "saldo_medio": c["saldo_medio"], "punteggio": c["punteggio"],
                         "ultima": c["ultima"], "stato": stato_c, "motivo": c["motivo"] or m.get("scartato") or m.get("uscita") or ""})
    for x in R["esclusi"]:
        cand_out.append({"ticker": x["ticker"], "nome": x["nome"], "isin": x["isin"], "parte": "-", "gruppo": "-", "D": None,
                         "W": None, "M": None, "saldo_medio": None, "punteggio": None, "ultima": None, "stato": "escluso", "motivo": x["motivo"]})
    sat = sum(w for t, w in pesi.items() if per_t.get(t) and per_t[t]["parte"] == "satellite")
    core = sum(w for t, w in pesi.items() if per_t.get(t) and per_t[t]["parte"] == "core")
    out = {"generato": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "data_esecuzione": oggi.isoformat(),
           "dati_al": rif, "regole_versione": R["versione"], "fonte": "Yahoo Finance", "azione": r["azione"], "motivo": r["motivo"],
           "stress": {"attivo": r["stress"], "quota": round(r["quota_stress"], 1), "soglia": P["stress_quota_peso"]},
           "quote": {"core": round(core, 1), "satellite": round(sat, 1), "paracadute": round(pesi.get(pc_t, 0.0), 1)},
           "nav": punto, "nav_serie": stato["nav_serie"], "posizioni": righe, "candidati": cand_out,
           "movimenti_ultimi": r["movimenti"], "storico": stato["storico"], "avvisi": avvisi,
           "benchmark": P["benchmark"], "paracadute": pc_t}
    return R, out, stato


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--forza-mensile", action="store_true"); ap.add_argument("--prova", action="store_true")
    ap.add_argument("--oggi", default=None); ap.add_argument("--cache", default=None)
    args = ap.parse_args()
    R, out, stato = esegui_live(args); errore = None
    if not args.prova:
        os.makedirs(DATA, exist_ok=True)
        with open(os.path.join(DATA, "stato.json"), "w", encoding="utf-8") as f:
            json.dump(stato, f, ensure_ascii=False, indent=1, allow_nan=False)
        with open(os.path.join(DATA, "portafoglio.json"), "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, allow_nan=False)
        try:
            import esporta
            esporta.scrivi_xlsx(out, R, os.path.join(DATA, "portafoglio_lgim.xlsx"))
            esporta.scrivi_pdf(R, os.path.join(DATA, "regole_lgim.pdf"))
        except Exception as e:  # lo stato e' gia' salvato: l'errore resta visibile ma non perde il lavoro
            errore = f"esportazione Excel/PDF non riuscita: {e!r}"
    print(f"{out['data_esecuzione']} dati al {out['dati_al']}: {out['azione']} - {out['motivo']}")
    for r in out["posizioni"]:
        print(f"  {r['ticker']:9s} {r['peso']:5.1f}% (target {r['target']:5.1f}) {r['parte']:10s} D={r['D']} W={r['W']} M={r['M']}")
    for a in out["avvisi"]:
        print("  AVVISO:", a)
    if errore:
        print("ERRORE:", errore); sys.exit(1)


if __name__ == "__main__":
    main()
