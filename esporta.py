"""Esportazione Excel (scarico) e PDF (regole) del portafoglio LGIM.

Il PDF e' generato da regole.json a ogni esecuzione: testo e parametri non possono divergere dal motore.
"""
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONT = "Arial"
INTEST = PatternFill("solid", start_color="1F3A5F")


def _foglio(wb, titolo, intestazioni, righe, larghezze, primo=False):
    ws = wb.active if primo else wb.create_sheet()
    ws.title = titolo
    ws.append(intestazioni)
    for r in righe:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); c.fill = INTEST
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=FONT, size=10)
            if isinstance(c.value, float):
                c.number_format = "0.0"
    for i, w in enumerate(larghezze, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def scrivi_xlsx(out, R, path):
    wb = Workbook()
    _foglio(wb, "Portafoglio",
            ["Ticker", "Nome", "ISIN", "Parte", "Area / famiglia", "Peso %", "Target %", "Saldo G", "Saldo S", "Saldo M",
             "Punteggio", "In portafoglio dal", "Dati"],
            [[r["ticker"], r["nome"], r["isin"], r["parte"], r["gruppo"], r["peso"], r["target"], r["D"], r["W"], r["M"],
              r["punteggio"], r["ingresso"], r["dati"]] for r in out["posizioni"]],
            [11, 46, 15, 11, 24, 9, 9, 9, 9, 9, 10, 16, 14], primo=True)
    _foglio(wb, "Candidati",
            ["Ticker", "Nome", "ISIN", "Parte", "Area / famiglia", "Saldo G", "Saldo S", "Saldo M", "Saldo medio", "Punteggio",
             "Ultima barra", "Stato", "Motivo"],
            [[c["ticker"], c["nome"], c["isin"], c["parte"], c["gruppo"], c["D"], c["W"], c["M"], c["saldo_medio"],
              c["punteggio"], c["ultima"], c["stato"], c["motivo"]] for c in out["candidati"]],
            [11, 46, 15, 11, 24, 9, 9, 9, 11, 10, 13, 26, 60])
    mov = []
    for s in out["storico"]:
        if not s["movimenti"]:
            mov.append([s["data"], s["dati_al"], s["azione"], s["motivo"], None, None, None, None, None, None])
        for m in s["movimenti"]:
            mov.append([s["data"], s["dati_al"], s["azione"], s["motivo"], m["ticker"], m["nome"], m["azione"],
                        m["da"], m["a"], m["motivo"]])
    _foglio(wb, "Movimenti",
            ["Data", "Dati al", "Azione", "Motivo dell'azione", "Ticker", "Nome", "Movimento", "Da %", "A %", "Motivo del movimento"],
            mov, [12, 12, 14, 50, 11, 40, 11, 8, 8, 44])
    bm = out["benchmark"]
    _foglio(wb, "Valore del modello", ["Data", "Dati al", "Azione", "Modello"] + bm,
            [[p["data"], p["dati_al"], p["azione"], p["nav"]] + [p.get(b) for b in bm] for p in out["nav_serie"]],
            [12, 12, 16, 11] + [12] * len(bm))
    P = R["parametri"]
    righe = [["Versione regole", R["versione"]], ["Aggiornate il", R["aggiornato"]],
             ["Parte core, peso obiettivo %", P["core_target"]],
             ["Parte satellite, peso obiettivo %", P["satellite"]["target"]],
             ["Banda satellite %", f"{P['satellite']['min']} - {P['satellite']['max']}"],
             ["Tetto per ETF %", P["tetto_peso"]], ["Soglia ribilanciamento settimanale (punti)", P["soglia_ribilanciamento"]],
             ["Stress: quota di peso con saldo G e S negativi %", P["stress_quota_peso"]],
             ["Costo per lato %", P["costo_per_lato_pct"]], ["Paracadute", P["paracadute"]],
             ["Benchmark", ", ".join(P["benchmark"])], [None, None]]
    righe += [[t, x] for t, x in R["testo"]]
    ws = _foglio(wb, "Regole", ["Voce", "Valore"], righe, [46, 120])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)


# ------------------------------------------------------------------------ PDF
def _sicuro(t):
    """Il PDF usa font standard: niente simboli fuori dal set Latin-1/WinAnsi."""
    for a, b in [("≥", ">="), ("≤", "<="), ("→", "->"), ("–", "-"), ("’", "'"), ("“", '"'), ("”", '"')]:
        t = t.replace(a, b)
    return t


def segnaposti(R):
    P = R["parametri"]; pc, ps = P["posizioni"]["core"], P["posizioni"]["satellite"]
    n_core = sum(1 for u in R["universo"] if u["parte"] == "core" and u.get("attivo", True))
    n_sat = sum(1 for u in R["universo"] if u["parte"] == "satellite" and u.get("attivo", True))
    n_ind = len(P["voti"]["indicatori"])
    n_aree = len({u["area"] for u in R["universo"] if u["parte"] == "core" and u.get("attivo", True)})
    return {"n_core": n_core, "n_aree": n_aree, "n_sat": n_sat, "max_area": P["max_per_area_core"], "max_fam": P["max_per_famiglia_satellite"],
            "n_voti_gs": n_ind + 2 * len(P["voti"]["medie_giornaliero_settimanale"]), "n_voti_m": n_ind + 2 * len(P["voti"]["medie_mensile"]),
            "core_min": pc["min"], "core_max": pc["max"], "core_in": pc["entra_entro"], "core_out": pc["esce_oltre"],
            "sat_pmin": ps["min"], "sat_pmax": ps["max"], "sat_in": ps["entra_entro"], "sat_out": ps["esce_oltre"],
            "perm": P["permanenza_minima_settimane"], "core_t": P["core_target"], "sat_t": P["satellite"]["target"],
            "sat_min": P["satellite"]["min"], "sat_max": P["satellite"]["max"], "tetto": P["tetto_peso"],
            "paracadute": P["paracadute"], "soglia": P["soglia_ribilanciamento"], "stress": P["stress_quota_peso"],
            "costo": str(P["costo_per_lato_pct"]).replace(".", ",")}


def testo_regole(R):
    s = segnaposti(R)
    return [(t, x.format(**s)) for t, x in R["testo"]]


def scrivi_pdf(R, path):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    st = getSampleStyleSheet()
    titolo = ParagraphStyle("t", parent=st["Title"], fontName="Helvetica-Bold", fontSize=18, alignment=0, spaceAfter=4)
    h = ParagraphStyle("h", parent=st["Heading2"], fontName="Helvetica-Bold", fontSize=12, spaceBefore=10, spaceAfter=3)
    b = ParagraphStyle("b", parent=st["BodyText"], fontName="Helvetica", fontSize=9.5, leading=13)
    pic = ParagraphStyle("p", parent=b, fontSize=8.5, textColor=colors.HexColor("#555555"))
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.8 * cm, bottomMargin=1.8 * cm,
                            title="Regole del portafoglio LGIM")
    story = [Paragraph("Regole del portafoglio LGIM", titolo),
             Paragraph(escape(_sicuro(f"Versione {R['versione']}, aggiornate il {R['aggiornato']}. Generato automaticamente da regole.json.")), pic)]
    for t, x in testo_regole(R):
        story += [Paragraph(escape(_sicuro(t)), h), Paragraph(escape(_sicuro(x)), b)]
    P = R["parametri"]
    tab = [["Parametro", "Valore"],
           ["Core / satellite (peso obiettivo)", f"{P['core_target']}% / {P['satellite']['target']}%"],
           ["Banda satellite", f"{P['satellite']['min']}% - {P['satellite']['max']}%"],
           ["Tetto per singolo ETF", f"{P['tetto_peso']}%"],
           ["Soglia ribilanciamento settimanale", f"{P['soglia_ribilanciamento']} punti"],
           ["Stress (quota di peso con saldo G e S negativi)", f"{P['stress_quota_peso']}%"],
           ["Permanenza minima", f"{P['permanenza_minima_settimane']} settimane"],
           ["Costo per lato nel valore del modello", f"{str(P['costo_per_lato_pct']).replace('.', ',')}%"],
           ["Paracadute", P["paracadute"]], ["Benchmark", ", ".join(P["benchmark"])]]
    t = Table(tab, colWidths=[9.5 * cm, 7.5 * cm])
    t.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                           ("FONTSIZE", (0, 0), (-1, -1), 9), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                           ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BBBBBB")),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story += [Paragraph("Parametri in vigore", h), t, Spacer(1, 8), Paragraph("Universo", h)]
    righe = [["Ticker", "Nome", "Parte", "Area / famiglia"]] + [[u["ticker"], _sicuro(u["nome"]), u["parte"], u.get("area") or u.get("famiglia")]
                                                                  for u in R["universo"]]
    tu = Table(righe, colWidths=[2.2 * cm, 8.6 * cm, 2.2 * cm, 4 * cm], repeatRows=1)
    tu.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                            ("FONTSIZE", (0, 0), (-1, -1), 8), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F6F9")])]))
    story.append(tu)
    doc.build(story)
