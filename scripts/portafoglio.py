#!/usr/bin/env python3
"""Portafoglio modello LGIM.

Legge i dati gia' calcolati da `investing` (checkout in ./investing o INVESTING_DIR),
applica regole.json e aggiorna data/stato.json, data/portafoglio.json, l'Excel e il PDF.

Uso:
  python scripts/portafoglio.py                  esecuzione normale
  python scripts/portafoglio.py --forza-mensile  forza il ribilanciamento mensile
  python scripts/portafoglio.py --prova          calcola e stampa, non scrive nulla
  python scripts/portafoglio.py --oggi 2026-11-02   data di esecuzione (solo per test)
"""
import argparse, collections, datetime as dt, json, os, sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INV = os.environ.get("INVESTING_DIR", os.path.join(BASE, "investing"))
DATA = os.environ.get("LGIM_DATA", os.path.join(BASE, "data"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


# ---------------------------------------------------------------- lettura dati
def carica_regole():
    return json.load(open(os.path.join(BASE, "regole.json"), encoding="utf-8"))


def carica_indice():
    p = os.path.join(INV, "data", "tech", "index.json")
    if not os.path.exists(p):
        sys.exit(f"ERRORE: indice di investing non trovato ({p})")
    return {r["ticker"]: r for r in json.load(open(p, encoding="utf-8"))["righe"]}


def carica_dettaglio(ticker):
    p = os.path.join(INV, "data", "tech", ticker.replace(".", "_") + ".json")
    if not os.path.exists(p):
        return None
    try:
        return json.load(open(p, encoding="utf-8"))
    except Exception:
        return None


def saldo_fisso(tf, indicatori, periodi):
    """Saldo con set fisso di voti. None se manca anche un solo voto."""
    if not tf:
        return None
    ind = {i["nome"]: i.get("azione") for i in tf.get("indicatori", [])}
    med = {m["periodo"]: m for m in tf.get("medie", [])}
    voti = [ind.get(n) for n in indicatori]
    for p in periodi:
        m = med.get(p)
        voti += [m.get("v_sma"), m.get("v_ema")] if m else [None, None]
    if any(v is None for v in voti):
        return None
    return round((voti.count("Compra") - voti.count("Vendi")) / len(voti) * 100, 1)


def media(valori):
    v = [x for x in valori if x is not None]
    return round(sum(v) / len(v), 1) if v else None


# ------------------------------------------------------------------ candidati
def costruisci_candidati(R, idx):
    P = R["parametri"]; V = P["voti"]
    out = []
    for u in R["universo"]:
        c = dict(u); t = u["ticker"]
        c["gruppo"] = u.get("area") or u.get("famiglia")
        c["stato"] = "ok"; c["motivo"] = None
        row = idx.get(t); det = carica_dettaglio(t) if row else None
        c["prezzo"] = row.get("prezzo") if row else None
        c["D"] = c["W"] = c["M"] = None; c["ultima"] = None
        c["saldo_medio"] = c["punteggio"] = None
        c["ok_ranking"] = False; c["ok_ingresso"] = False; c["vecchio"] = False
        if not row:
            c["stato"], c["motivo"] = "assente", "non presente nei dati di investing (storia insufficiente o ticker non coperto)"
        elif not det:
            c["stato"], c["motivo"] = "senza_dettaglio", "dettaglio non disponibile: aggiungere il ticker a campione.txt di investing"
        else:
            tf = det.get("tf") or {}
            c["D"] = saldo_fisso(tf.get("D"), V["indicatori"], V["medie_giornaliero_settimanale"])
            c["W"] = saldo_fisso(tf.get("W"), V["indicatori"], V["medie_giornaliero_settimanale"])
            c["M"] = saldo_fisso(tf.get("M"), V["indicatori"], V["medie_mensile"])
            c["ultima"] = (tf.get("D") or {}).get("ultima")
            if det.get("prezzo") is not None:
                c["prezzo"] = det["prezzo"]
            if c["D"] is None or c["W"] is None:
                c["stato"], c["motivo"] = "storia_insufficiente", "storia insufficiente per il set fisso di voti (giornaliero o settimanale)"
            else:
                c["saldo_medio"] = media([c["D"], c["W"], c["M"]])
                c["punteggio"] = media([c["W"], c["M"]])
                c["ok_ranking"] = c["saldo_medio"] > 0
                if not c["ok_ranking"]:
                    c["motivo"] = "saldo medio non positivo"
        out.append(c)
    # data di riferimento: la piu' frequente tra gli ETF con dati (a parita' la piu' recente)
    date = collections.Counter(c["ultima"] for c in out if c["ultima"])
    rif = max(date.items(), key=lambda kv: (kv[1], kv[0]))[0] if date else None
    for c in out:
        c["vecchio"] = bool(rif and c["ultima"] and c["ultima"] < rif)
        c["ok_ingresso"] = bool(c["ok_ranking"] and (c["D"] or 0) > 0 and not c["vecchio"])
        if c["vecchio"] and c["motivo"] is None:
            c["motivo"] = f"ultima barra {c['ultima']} piu' vecchia del riferimento {rif}"
    return out, rif


# ------------------------------------------------------------------ selezione
def scegli_parte(parte, cands, tenuti, P, oggi, ingressi, congelati):
    """Restituisce (posizioni scelte, meta con i motivi)."""
    pc = P["posizioni"][parte]; chiave = "area" if parte == "core" else "famiglia"
    maxg = P["max_per_area_core"] if parte == "core" else P["max_per_famiglia_satellite"]
    elig = [c for c in cands if c["parte"] == parte and c["ok_ranking"]]
    elig.sort(key=lambda c: (-c["punteggio"], -c["saldo_medio"], c["ticker"]))
    cont = collections.Counter(); lista = []; meta = {}
    for c in elig:
        g = c[chiave]
        if cont[g] >= maxg:
            meta[c["ticker"]] = {"scartato": f"limite di {maxg} per {g}"}; continue
        cont[g] += 1; lista.append(c)
    rank = {c["ticker"]: i + 1 for i, c in enumerate(lista)}
    per_t = {c["ticker"]: c for c in cands}
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
        if c["vecchio"]:
            pos.append(t); meta[t] = {"nota": "dati non aggiornati: resta fermo"}; continue
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
    cont_pos = collections.Counter(per_t[t][chiave] for t in pos if t in per_t)
    for c in lista:
        if len(pos) >= pc["max"]:
            break
        if cont_pos[c[chiave]] >= maxg and c["ticker"] not in pos:
            continue
        if c["ticker"] not in pos and c["ok_ingresso"] and rank[c["ticker"]] <= pc["entra_entro"]:
            cont_pos[c[chiave]] += 1
            pos.append(c["ticker"]); meta[c["ticker"]] = {"ingresso": f"posizione {rank[c['ticker']]} nella parte {parte}"}
    for c in lista:
        if len(pos) >= pc["min"]:
            break
        if cont_pos[c[chiave]] >= maxg and c["ticker"] not in pos:
            continue
        if c["ticker"] not in pos and c["ok_ingresso"]:
            cont_pos[c[chiave]] += 1
            pos.append(c["ticker"]); meta[c["ticker"]] = {"ingresso": f"posizione {rank[c['ticker']]}, per raggiungere il minimo di {pc['min']}"}
    return pos, meta


def pesi_parte(parte, pos, cands, P, congelati, cw):
    per_t = {c["ticker"]: c for c in cands}; pc = P["posizioni"][parte]
    budget = P["core_target"] if parte == "core" else P["satellite"]["target"]
    fissi = {t: cw.get(t, 0.0) for t in pos if t in congelati}
    liberi = [t for t in pos if t not in congelati]
    budget = max(0.0, budget - sum(fissi.values()))
    frazione = min(1.0, len(pos) / pc["min"])
    B = budget * frazione
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


def obiettivo(cands, tenuti, P, oggi, ingressi, cw):
    per_t = {c["ticker"]: c for c in cands}
    congelati = {t for t in tenuti if t in per_t and (per_t[t]["punteggio"] is None or per_t[t]["vecchio"])}
    tgt = {}; meta = {}
    for parte in ("core", "satellite"):
        pos, m = scegli_parte(parte, cands, tenuti, P, oggi, ingressi, congelati)
        meta.update(m); tgt.update(pesi_parte(parte, pos, cands, P, congelati, cw))
    tgt = {t: round(v, 2) for t, v in tgt.items() if v > 0.004}
    tgt[P["paracadute"]] = round(max(0.0, 100.0 - sum(tgt.values())), 2)
    return tgt, meta


# ---------------------------------------------------------------------- stato
def leggi_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def prezzo_di(t, idx, stato):
    r = idx.get(t)
    if r and r.get("prezzo") is not None:
        return float(r["prezzo"])
    return stato.get("posizioni", {}).get(t, {}).get("prezzo")


def esegui(args):
    R = carica_regole(); P = R["parametri"]; idx = carica_indice()
    oggi = dt.date.fromisoformat(args.oggi) if args.oggi else dt.datetime.now(dt.timezone.utc).date()
    cands, rif = costruisci_candidati(R, idx)
    per_t = {c["ticker"]: c for c in cands}
    pc_t = P["paracadute"]
    sp = os.path.join(DATA, "stato.json")
    stato = leggi_json(sp, None); nuovo = stato is None
    avvisi = []
    if nuovo:
        stato = {"versione": 1, "creato": oggi.isoformat(), "valore": 100.0, "posizioni": {},
                 "ultimo_mensile": None, "benchmark_base": {}, "nav_serie": [], "storico": []}
    # valorizzazione corrente
    val = {}; prezzi = {}
    for t, pz in stato["posizioni"].items():
        p = prezzo_di(t, idx, stato)
        if p is None:
            avvisi.append(f"Prezzo non disponibile per {t}: usato l'ultimo noto")
            p = pz["prezzo"]
        prezzi[t] = p; val[t] = pz["unita"] * p
    totale = sum(val.values()) if val else stato["valore"]
    cw = {t: v / totale * 100 for t, v in val.items()} if val else {}
    tenuti = [t for t in cw if t != pc_t and cw[t] > 0]
    ingressi = {t: pz.get("ingresso", oggi.isoformat()) for t, pz in stato["posizioni"].items()}
    # stress
    eq = sum(cw[t] for t in tenuti)
    negativi = [t for t in tenuti if per_t.get(t) and per_t[t]["D"] is not None and per_t[t]["W"] is not None
                and per_t[t]["D"] < 0 and per_t[t]["W"] < 0 and not per_t[t]["vecchio"]]
    quota_stress = (sum(cw[t] for t in negativi) / eq * 100) if eq > 0 else 0.0
    stress = eq > 0 and quota_stress >= P["stress_quota_peso"]
    tgt, meta = obiettivo(cands, tenuti, P, oggi, ingressi, cw)
    mese = oggi.strftime("%Y-%m")
    # decisione
    if nuovo:
        azione, motivo, nw = "INIZIO", "prima esecuzione: il portafoglio adotta il target", dict(tgt)
    elif stress:
        azione = "STRESS"
        motivo = f"{quota_stress:.0f}% del peso azionario con saldo giornaliero e settimanale negativi"
        nw = {t: (0.0 if t in negativi else cw[t]) for t in tenuti}
        nw[pc_t] = 100.0 - sum(nw.values())
        for t in negativi:
            meta.setdefault(t, {})["uscita"] = "stress: saldo giornaliero e settimanale negativi"
    elif args.forza_mensile or stato.get("ultimo_mensile") != mese:
        azione, motivo, nw = "MENSILE", "ribilanciamento mensile al target", dict(tgt)
    else:
        sat_w = sum(cw.get(t, 0) for t in tenuti if per_t.get(t) and per_t[t]["parte"] == "satellite")
        fuori = eq > 0 and not (P["satellite"]["min"] <= sat_w <= P["satellite"]["max"]) and sat_w > 0
        dev = {t: abs(tgt.get(t, 0.0) - cw.get(t, 0.0)) for t in set(tgt) | set(tenuti) if t != pc_t}
        sopra = [t for t, d in dev.items() if d >= P["soglia_ribilanciamento"]]
        if fuori:
            azione, motivo, nw = "SETTIMANALE", f"satellite al {sat_w:.1f}%, fuori dalla banda {P['satellite']['min']}-{P['satellite']['max']}%", dict(tgt)
        elif sopra:
            azione, motivo = "SETTIMANALE", "scostamento di almeno %g punti su %s" % (P["soglia_ribilanciamento"], ", ".join(sorted(sopra)))
            nw = {t: (tgt.get(t, 0.0) if t in sopra else cw.get(t, 0.0)) for t in set(tgt) | set(tenuti) if t != pc_t}
            eccesso = sum(nw.values()) - 100.0
            if eccesso > 0:
                aumenti = {t: nw[t] - cw.get(t, 0.0) for t in nw if nw[t] > cw.get(t, 0.0)}
                tot_a = sum(aumenti.values())
                for t, a in aumenti.items():
                    nw[t] -= eccesso * a / tot_a
            nw[pc_t] = max(0.0, 100.0 - sum(nw.values()))
        else:
            azione, motivo, nw = "MANTENIMENTO", "nessuno scostamento rilevante", dict(cw)
    nw = {t: round(v, 4) for t, v in nw.items() if v > 0.004}
    # movimenti, costi, unita'
    movimenti = []
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
            mot = m.get("ingresso") or m.get("uscita") or (azione.lower() if t != pc_t else "quota non assegnata" if tipo in ("ENTRA", "AUMENTA") else "riallocazione")
            movimenti.append({"ticker": t, "nome": (per_t.get(t) or {}).get("nome") or (idx.get(t) or {}).get("nome", t),
                              "azione": tipo, "da": round(a, 1), "a": round(b, 1), "motivo": mot})
        posizioni = {}
        for t, w in nw.items():
            p = prezzi.get(t) or prezzo_di(t, idx, stato)
            if p is None:
                avvisi.append(f"Prezzo non disponibile per {t}: posizione non aperta"); continue
            vecchio = stato["posizioni"].get(t)
            posizioni[t] = {"unita": dopo * w / 100 / p, "prezzo": p,
                            "ingresso": vecchio["ingresso"] if vecchio else oggi.isoformat()}
        stato["posizioni"] = posizioni; stato["valore"] = dopo
        if azione in ("INIZIO", "MENSILE") or nuovo:
            stato["ultimo_mensile"] = mese
        pesi_fin = nw; valore_fin = dopo
    else:
        for t, pz in stato["posizioni"].items():
            pz["prezzo"] = prezzi[t]
        stato["valore"] = totale; pesi_fin = cw; valore_fin = totale
    # benchmark e NAV
    bm = {}
    for b in P["benchmark"]:
        p = prezzo_di(b, idx, {})
        if p is None:
            avvisi.append(f"Prezzo del benchmark {b} non disponibile"); continue
        stato["benchmark_base"].setdefault(b, p); bm[b] = round(p / stato["benchmark_base"][b] * 100, 2)
    punto = {"data": oggi.isoformat(), "dati_al": rif, "nav": round(valore_fin, 2), **bm, "azione": azione}
    serie = [x for x in stato["nav_serie"] if x["data"] != punto["data"]] + [punto]
    stato["nav_serie"] = serie[-400:]
    if azione not in ("MANTENIMENTO",):
        stato["storico"] = ([{"data": oggi.isoformat(), "dati_al": rif, "azione": azione, "motivo": motivo,
                              "movimenti": movimenti}] + stato["storico"])[:200]
    stato["ultima_esecuzione"] = oggi.isoformat()
    # avvisi
    mancanti = [c["ticker"] for c in cands if c["stato"] in ("assente", "senza_dettaglio")]
    if mancanti:
        avvisi.append("ETF senza dati utilizzabili su investing (%d): %s" % (len(mancanti), ", ".join(mancanti)))
    vecchi = [c["ticker"] for c in cands if c["vecchio"]]
    if vecchi:
        avvisi.append("ETF con barra piu' vecchia del riferimento %s: %s" % (rif, ", ".join(vecchi)))
    # uscita per la pagina
    righe = []
    for t, w in sorted(pesi_fin.items(), key=lambda kv: -kv[1]):
        c = per_t.get(t)
        r = idx.get(t, {})
        righe.append({"ticker": t, "nome": (c or {}).get("nome") or r.get("nome", t), "isin": (c or {}).get("isin") or r.get("isin"),
                      "parte": (c or {}).get("parte", "paracadute"), "gruppo": (c or {}).get("gruppo"),
                      "peso": round(w, 1), "target": tgt.get(t, 0.0), "D": (c or {}).get("D"), "W": (c or {}).get("W"),
                      "M": (c or {}).get("M"), "punteggio": (c or {}).get("punteggio"),
                      "ingresso": stato["posizioni"].get(t, {}).get("ingresso"), "prezzo": prezzi.get(t) or r.get("prezzo"),
                      "dati": "aggiornati" if c and not c["vecchio"] and c["ultima"] else ("non aggiornati" if c and c["vecchio"] else "n.d.")})
    cand_out = []
    for c in cands:
        m = meta.get(c["ticker"], {})
        in_p = c["ticker"] in pesi_fin and c["ticker"] != pc_t
        stato_c = "in portafoglio" if in_p else ("escluso" if c["stato"] != "ok" else
                  "qualificato, non selezionato" if c["ok_ranking"] else "non qualificato")
        cand_out.append({"ticker": c["ticker"], "nome": c["nome"], "isin": c["isin"], "parte": c["parte"], "gruppo": c["gruppo"],
                         "D": c["D"], "W": c["W"], "M": c["M"], "saldo_medio": c["saldo_medio"], "punteggio": c["punteggio"],
                         "ultima": c["ultima"], "stato": stato_c,
                         "motivo": c["motivo"] or m.get("scartato") or m.get("uscita") or ""})
    for x in R["esclusi"]:
        cand_out.append({"ticker": x["ticker"], "nome": x["nome"], "isin": x["isin"], "parte": "-", "gruppo": "-", "D": None,
                         "W": None, "M": None, "saldo_medio": None, "punteggio": None, "ultima": None, "stato": "escluso", "motivo": x["motivo"]})
    sat_fin = sum(w for t, w in pesi_fin.items() if per_t.get(t) and per_t[t]["parte"] == "satellite")
    core_fin = sum(w for t, w in pesi_fin.items() if per_t.get(t) and per_t[t]["parte"] == "core")
    out = {"generato": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "data_esecuzione": oggi.isoformat(), "dati_al": rif,
           "regole_versione": R["versione"], "azione": azione, "motivo": motivo,
           "stress": {"attivo": bool(stress), "quota": round(quota_stress, 1), "soglia": P["stress_quota_peso"]},
           "quote": {"core": round(core_fin, 1), "satellite": round(sat_fin, 1), "paracadute": round(pesi_fin.get(pc_t, 0.0), 1)},
           "nav": punto, "nav_serie": stato["nav_serie"], "posizioni": righe, "candidati": cand_out,
           "movimenti_ultimi": movimenti, "storico": stato["storico"], "avvisi": avvisi,
           "benchmark": P["benchmark"], "paracadute": pc_t}
    return R, out, stato


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--forza-mensile", action="store_true"); ap.add_argument("--prova", action="store_true")
    ap.add_argument("--oggi", default=None)
    args = ap.parse_args()
    R, out, stato = esegui(args)
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
        except ImportError as e:
            print("AVVISO: esportazione Excel/PDF non eseguita:", e)

    print(f"{out['data_esecuzione']} dati al {out['dati_al']}: {out['azione']} - {out['motivo']}")
    for r in out["posizioni"]:
        print(f"  {r['ticker']:9s} {r['peso']:5.1f}% (target {r['target']:5.1f}) {r['parte']:10s} D={r['D']} W={r['W']} M={r['M']}")
    for a in out["avvisi"]:
        print("  AVVISO:", a)

if __name__ == "__main__":
    main()
