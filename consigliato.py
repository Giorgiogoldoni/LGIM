"""Consigliato LGIM: pesi suggeriti da L&G (statici, solo visivi) piu' una variazione mensile cumulativa
indicata dai nostri indicatori. Nessun valore del modello, nessuna operazione.

Tutti i parametri stanno in regole.json (chiavi "consigliato" e "testo_consigliato"): il PDF delle regole
e il motore leggono lo stesso file e non possono divergere.
"""
import datetime as dt
import math

MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre",
        "novembre", "dicembre"]


def _arrotonda(x):
    """Arrotondamento all'intero simmetrico rispetto allo zero (0,5 va a 1 e -0,5 a -1)."""
    return int(math.copysign(math.floor(abs(x) + 0.5), x))


def _scala(d, obiettivo):
    """Riduce i valori positivi di d in proporzione fino a sommare esattamente `obiettivo` (interi, resti maggiori)."""
    tot = sum(d.values())
    quote = {t: x * obiettivo / tot for t, x in d.items()}
    base = {t: int(q) for t, q in quote.items()}
    resto = obiettivo - sum(base.values())
    ordine = sorted(d, key=lambda t: (-(quote[t] - base[t]), -d[t], t))
    for t in ordine[:resto]:
        base[t] += 1
    return base


def bilancia(var, minimo):
    """Rende la somma delle variazioni uguale a zero: il lato piu' grande viene ridotto fino a pareggiare l'altro.
    Se dopo la riduzione una variazione scende sotto `minimo` punti viene azzerata e il calcolo si ripete."""
    v = dict(var)
    for _ in range(len(v) + 2):
        pos = {t: x for t, x in v.items() if x > 0}
        neg = {t: -x for t, x in v.items() if x < 0}
        P, N = sum(pos.values()), sum(neg.values())
        if P == 0 or N == 0:
            return {t: 0 for t in v}
        if P > N:
            pos = _scala(pos, N)
        elif N > P:
            neg = _scala(neg, P)
        nuovo = {t: 0 for t in v}
        nuovo.update(pos)
        nuovo.update({t: -x for t, x in neg.items()})
        piccoli = [t for t, x in nuovo.items() if 0 < abs(x) < minimo]
        if not piccoli:
            return nuovo
        for t in piccoli:
            nuovo[t] = 0
        v = nuovo
    return {t: 0 for t in v}


def variazioni(pesi, punteggi, V, tetto):
    """Variazione in punti per ogni ETF, a somma zero. punteggi: {ticker: punteggio oppure None}.
    Ritorna (variazioni, media_ponderata oppure None)."""
    val = {t: punteggi[t] for t in pesi if punteggi.get(t) is not None}
    peso_val = sum(pesi[t] for t in val)
    if len(val) < 2 or peso_val <= 0:
        return {t: 0 for t in pesi}, None
    media = sum(pesi[t] * val[t] for t in val) / peso_val
    mn, mx, k = V["min_punti"], V["max_punti"], V["coefficiente"]
    grezze = {}
    for t in pesi:
        if t not in val:
            grezze[t] = 0
            continue
        w = pesi[t]
        v = _arrotonda((val[t] - media) * k)
        giu = min(mx, w)
        su = max(0, min(mx, max(w, mn), tetto - w))
        v = max(-giu, min(su, v))
        grezze[t] = v if abs(v) >= mn else 0
    return bilancia(grezze, mn), media


def pesi_base(R):
    C = R["consigliato"]
    base = {t: int(round(w)) for t, w in C["pesi"].items()}
    if sum(base.values()) != 100 or any(abs(C["pesi"][t] - base[t]) > 1e-9 for t in base):
        raise ValueError("i pesi del consigliato in regole.json devono essere interi con somma 100")
    return base


def _mese_it(iso_mese):
    a, m = iso_mese.split("-")
    return f"{MESI[int(m) - 1]} {a}"


def _mese_dopo(iso_mese):
    a, m = (int(x) for x in iso_mese.split("-"))
    a, m = (a + 1, 1) if m == 12 else (a, m + 1)
    return f"{a:04d}-{m:02d}"


def aggiorna(R, out, stato, forza=False):
    """out: contenuto di portafoglio.json (candidati con punteggio, dati_al, data_esecuzione).
    stato: contenuto precedente di consigliato_stato.json oppure None.
    Ritorna (consigliato.json, nuovo stato). Funzione pura: non scrive nulla."""
    C, T = R["consigliato"], R["parametri"]
    V, tetto = C["variazione"], T["tetto_peso"]
    base = pesi_base(R)
    oggi = out["data_esecuzione"]
    mese = oggi[:7]
    rif = out.get("dati_al")
    per_t = {c["ticker"]: c for c in out["candidati"]}
    univ = {u["ticker"]: u for u in R["universo"]}
    avvisi, punteggi, non_val = [], {}, []
    for t in base:
        c = per_t.get(t)
        vecchio = bool(c and c.get("ultima") and rif and c["ultima"] < rif)
        if c and c.get("punteggio") is not None and not vecchio:
            punteggi[t] = c["punteggio"]
        else:
            punteggi[t] = None
            non_val.append(t)
    stato = dict(stato) if stato else None
    nuova_versione = stato is None or stato.get("valido_dal") != C["valido_dal"]
    stesso_mese = (not nuova_versione) and stato.get("mese") == mese
    if nuova_versione:
        stato = {"versione": 1, "valido_dal": C["valido_dal"], "mese": None, "aggiornato_il": None,
                 "pesi": dict(base), "pesi_prima": dict(base), "storico": (stato or {}).get("storico", [])}
        motivo = "nuova versione L&G: partenza dai pesi L&G" if (stato["storico"]) else "primo calcolo: partenza dai pesi L&G"
    else:
        motivo = "aggiornamento mensile"
    if nuova_versione or forza or not stesso_mese:
        if stesso_mese:  # ricalcolo forzato nello stesso mese: riparte dai pesi di inizio mese, senza sommare due volte
            partenza = dict(stato["pesi_prima"])
            stato["storico"] = [x for x in stato["storico"] if x.get("mese") != mese]
            motivo = "ricalcolo forzato dello stesso mese"
        elif nuova_versione:
            partenza = dict(base)
        else:
            partenza = dict(stato["pesi"])
        var, media = variazioni(partenza, punteggi, V, tetto)
        nuovi = {t: partenza[t] + var[t] for t in partenza}
        if sum(nuovi.values()) != 100 or any(w < 0 or w > tetto for w in nuovi.values()):
            avvisi.append("Il calcolo della variazione non ha rispettato i vincoli: pesi lasciati invariati")
            var, nuovi = {t: 0 for t in partenza}, dict(partenza)
        if media is None:
            avvisi.append("Meno di due ETF valutabili: nessuna variazione questo mese")
        stato["pesi_prima"] = partenza
        stato["pesi"] = nuovi
        stato["mese"] = mese
        stato["aggiornato_il"] = oggi
        voce = {"data": oggi, "mese": mese, "dati_al": rif, "motivo": motivo,
                "media_punteggi": None if media is None else round(media, 1),
                "variazioni": {t: x for t, x in var.items() if x}, "pesi": dict(nuovi)}
        stato["storico"] = ([voce] + stato["storico"])[:60]
    pesi = stato["pesi"]
    ultima = stato["storico"][0]["variazioni"] if stato["storico"] else {}
    righe = []
    for t in sorted(pesi, key=lambda x: (-pesi[x], x)):
        u, c = univ.get(t, {}), per_t.get(t, {})
        righe.append({"ticker": t, "nome": u.get("nome") or t, "isin": u.get("isin"),
                      "gruppo": u.get("area") or u.get("tema") or u.get("famiglia"),
                      "peso_lgim": base[t], "peso": pesi[t], "scarto": pesi[t] - base[t], "ultima_variazione": ultima.get(t, 0),
                      "D": c.get("D"), "W": c.get("W"), "M": c.get("M"), "punteggio": c.get("punteggio"),
                      "dati": "non valutabile" if t in non_val else "valutabile"})
    if non_val:
        avvisi.append("ETF non valutabili (senza dati o barra non aggiornata): " + ", ".join(non_val))
    try:
        scad = dt.date.fromisoformat(C["valido_dal"]) + dt.timedelta(days=183)
        if dt.date.fromisoformat(oggi) > scad:
            avvisi.append(f"I pesi L&G sono in vigore dal {C['valido_dal']}: sono passati più di 6 mesi, da aggiornare")
    except ValueError:
        avvisi.append("Data di validità dei pesi L&G non leggibile")
    cons = {"generato": out["generato"], "data_esecuzione": oggi, "dati_al": rif, "regole_versione": R.get("versione_consigliato", "1.0"),
            "fonte": C["fonte"], "valido_dal": C["valido_dal"], "mese": stato["mese"], "aggiornato_il": stato["aggiornato_il"],
            "prossimo_aggiornamento": _mese_it(_mese_dopo(stato["mese"])), "variazione": V, "tetto": tetto,
            "media_punteggi": stato["storico"][0]["media_punteggi"] if stato["storico"] else None,
            "posizioni": righe, "storico": stato["storico"], "avvisi": avvisi}
    return cons, stato
