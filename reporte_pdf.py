# -*- coding: utf-8 -*-
"""
Informe de auditoría en PDF (A4 horizontal): resumen ejecutivo, una sección
por año con gráficos, hallazgos, faltantes, repetidos y observaciones.
"""
import datetime
import math
import os

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
                                Table, TableStyle, PageBreak, KeepTogether, CondPageBreak)
from reportlab.graphics.shapes import Drawing, String, Rect, Line
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie

from auditoria_core import DESCRIPCION

# ---------------------------------------------------------------- estilo
AZUL = colors.HexColor("#1F4E78")
AZUL_CL = colors.HexColor("#DDEBF7")
GRIS = colors.HexColor("#595959")
GRIS_CL = colors.HexColor("#F2F2F2")
ROJO = colors.HexColor("#C0392B")
ROJO_CL = colors.HexColor("#FADBD8")
AMBAR = colors.HexColor("#E0A800")
AMBAR_TXT = colors.HexColor("#9A6B00")
AMBAR_CL = colors.HexColor("#FFF4CC")
VERDE = colors.HexColor("#2E8B57")
VERDE_CL = colors.HexColor("#D5F0E0")

PAGE = landscape(A4)
MARGEN = 15 * mm
ANCHO = PAGE[0] - 2 * MARGEN

E = {
    "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=26, leading=31, textColor=AZUL),
    "subtit": ParagraphStyle("subtit", fontName="Helvetica", fontSize=14, leading=18, textColor=GRIS),
    "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=AZUL, spaceAfter=6),
    "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=11.5, leading=14, textColor=AZUL,
                         spaceBefore=8, spaceAfter=4),
    "txt": ParagraphStyle("txt", fontName="Helvetica", fontSize=9.5, leading=13, alignment=TA_JUSTIFY),
    "txt_c": ParagraphStyle("txt_c", fontName="Helvetica", fontSize=9, leading=12, alignment=TA_CENTER),
    "peq": ParagraphStyle("peq", fontName="Helvetica", fontSize=8, leading=10),
    "num": ParagraphStyle("num", fontName="Courier", fontSize=8.5, leading=11),
    "celda": ParagraphStyle("celda", fontName="Helvetica", fontSize=8, leading=10),
    "kpi_v": ParagraphStyle("kpi_v", fontName="Helvetica-Bold", fontSize=17, leading=20, textColor=AZUL,
                            alignment=TA_CENTER),
    "kpi_t": ParagraphStyle("kpi_t", fontName="Helvetica", fontSize=7.5, leading=9, textColor=GRIS,
                            alignment=TA_CENTER),
}
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def fecha_larga(d=None):
    d = d or datetime.date.today()
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def miles(n):
    return f"{n:,}".replace(",", " ")


def pct(x):
    return "—" if x is None else f"{x * 100:.1f}%".replace(".", ",")


def rangos(nums):
    """[1,2,3,5,7,8] -> '1–3, 5, 7–8'"""
    nums = sorted(nums)
    if not nums:
        return ""
    out, ini, ant = [], nums[0], nums[0]
    for n in nums[1:] + [None]:
        if n is not None and n == ant + 1:
            ant = n
            continue
        out.append(str(ini) if ini == ant else (f"{ini}, {ant}" if ant == ini + 1 else f"{ini}–{ant}"))
        if n is not None:
            ini = ant = n
    return ", ".join(out)


def riesgo(comp):
    if comp is None:
        return "—", GRIS, GRIS_CL
    if comp >= 0.98:
        return "BAJO", VERDE, VERDE_CL
    if comp >= 0.90:
        return "MEDIO", AMBAR_TXT, AMBAR_CL
    return "ALTO", ROJO, ROJO_CL


def esc(t):
    return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# ---------------------------------------------------------------- lienzo con "Página X de Y"
def hacer_canvas(meta):
    class NumCanvas(rl_canvas.Canvas):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self._pags = []

        def showPage(self):
            self._pags.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._pags)
            for st in self._pags:
                self.__dict__.update(st)
                if self._pageNumber > 1:
                    self._decorar(total)
                super().showPage()
            super().save()

        def _decorar(self, total):
            w, h = PAGE
            self.setStrokeColor(AZUL)
            self.setLineWidth(1.2)
            self.line(MARGEN, h - 11 * mm, w - MARGEN, h - 11 * mm)
            self.setFont("Helvetica-Bold", 8)
            self.setFillColor(AZUL)
            self.drawString(MARGEN, h - 9.5 * mm, meta["entidad"] or "Informe de auditoría")
            self.setFont("Helvetica", 8)
            self.setFillColor(GRIS)
            self.drawRightString(w - MARGEN, h - 9.5 * mm,
                                 f"Auditoría de correlativos – {meta['tipo']}")
            self.setStrokeColor(colors.HexColor("#BFBFBF"))
            self.setLineWidth(0.5)
            self.line(MARGEN, 10 * mm, w - MARGEN, 10 * mm)
            self.drawString(MARGEN, 6.5 * mm, f"Fuente: {meta['archivo']}   |   Emitido el {meta['fecha']}")
            self.drawRightString(w - MARGEN, 6.5 * mm, f"Página {self._pageNumber} de {total}")
    return NumCanvas


# ---------------------------------------------------------------- gráficos
def _barras(valores, etiquetas, ancho, alto, titulo, color=ROJO, formato="%d", max_y=None, rot=0):
    d = Drawing(ancho, alto)
    d.add(String(4, alto - 12, titulo, fontName="Helvetica-Bold", fontSize=9, fillColor=AZUL))
    bc = VerticalBarChart()
    bc.x, bc.y = 34, 30 if rot else 22
    bc.width, bc.height = ancho - 44, alto - (58 if rot else 48)
    bc.data = [valores or [0]]
    bc.categoryAxis.categoryNames = etiquetas or [""]
    bc.categoryAxis.labels.fontName = "Helvetica"
    bc.categoryAxis.labels.fontSize = 6.5
    if rot:
        bc.categoryAxis.labels.angle = rot
        bc.categoryAxis.labels.boxAnchor = "ne"
        bc.categoryAxis.labels.dy = -2
    bc.valueAxis.labels.fontName = "Helvetica"
    bc.valueAxis.labels.fontSize = 7
    bc.valueAxis.valueMin = 0
    top = max_y or max(max(valores or [0]), 1)
    paso = 1
    for exp in range(0, 7):
        for m in (1, 2, 5):
            if top / (m * 10 ** exp) <= 5:
                paso = m * 10 ** exp
                break
        else:
            continue
        break
    if max_y:
        paso = max_y / 5
    bc.valueAxis.valueMax = max_y or paso * math.ceil(top * 1.12 / paso)
    bc.valueAxis.valueStep = paso
    bc.valueAxis.gridStrokeColor = colors.HexColor("#E0E0E0")
    bc.valueAxis.visibleGrid = True
    bc.bars[0].fillColor = color
    bc.bars[0].strokeColor = None
    bc.barLabels.fontName = "Helvetica"
    bc.barLabels.fontSize = 6.5
    bc.barLabelFormat = formato
    bc.barLabels.nudge = 6
    bc.barSpacing = 1
    d.add(bc)
    return d


def _torta(ok, rep, falt, ancho, alto):
    d = Drawing(ancho, alto)
    d.add(String(4, alto - 12, "Estado de la secuencia", fontName="Helvetica-Bold", fontSize=9, fillColor=AZUL))
    total = ok + rep + falt
    datos = [(ok, "Confirmados", VERDE), (rep, "Repetidos", AMBAR), (falt, "Faltantes", ROJO)]
    datos = [x for x in datos if x[0] > 0] or [(1, "Sin datos", GRIS)]
    p = Pie()
    diam = min(alto - 40, ancho * 0.48)
    p.x, p.y = 10, (alto - 18 - diam) / 2
    p.width = p.height = diam
    p.data = [x[0] for x in datos]
    p.labels = None
    p.slices.strokeColor = colors.white
    p.slices.strokeWidth = 1
    for i, x in enumerate(datos):
        p.slices[i].fillColor = x[2]
    d.add(p)
    lx, ly = p.x + diam + 14, alto / 2 + 18
    for i, (v, nom, col) in enumerate([(ok, "Números confirmados", VERDE), (rep, "Números repetidos", AMBAR),
                                       (falt, "Números faltantes", ROJO)]):
        y = ly - i * 20
        d.add(Rect(lx, y, 9, 9, fillColor=col, strokeColor=None))
        porc = f"{v / total * 100:.1f}%".replace(".", ",") if total else "0%"
        d.add(String(lx + 14, y + 1, f"{nom}: {miles(v)} ({porc})", fontName="Helvetica", fontSize=8))
    return d


def _tramos(seq, ultimo, n_tramos=20):
    """Faltantes agrupados por tramos de numeración."""
    if not ultimo:
        return [], []
    tam = max(5, int(math.ceil(ultimo / n_tramos / 5.0) * 5))
    falt = [f[0] for f in seq if f[2] == "FALTA"]
    cats, vals = [], []
    for ini in range(1, ultimo + 1, tam):
        fin = min(ini + tam - 1, ultimo)
        cats.append(f"{ini}-{fin}")
        vals.append(sum(1 for n in falt if ini <= n <= fin))
    return vals, cats


def _barra_progreso(comp, ancho=150, alto=12):
    d = Drawing(ancho, alto)
    _, col, _ = riesgo(comp)
    d.add(Rect(0, 0, ancho, alto, fillColor=GRIS_CL, strokeColor=None))
    d.add(Rect(0, 0, ancho * (comp or 0), alto, fillColor=col, strokeColor=None))
    return d


# ---------------------------------------------------------------- bloques
def _kpis(items, ancho=ANCHO):
    celdas = [[Paragraph(v, E["kpi_v"]) for _, v in items], [Paragraph(t, E["kpi_t"]) for t, _ in items]]
    t = Table(celdas, colWidths=[ancho / len(items)] * len(items))
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AZUL_CL),
        ("LINEAFTER", (0, 0), (-2, -1), 2, colors.white),
        ("TOPPADDING", (0, 0), (-1, 0), 7), ("BOTTOMPADDING", (0, 1), (-1, 1), 7),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def _tabla(datos, anchos, encabezado=True, zebra=True, alinear_centro_desde=0, extra=None):
    t = Table(datos, colWidths=anchos, repeatRows=1 if encabezado else 0)
    st = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if encabezado:
        st += [("BACKGROUND", (0, 0), (-1, 0), AZUL), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("ALIGN", (0, 0), (-1, 0), "CENTER")]
    if alinear_centro_desde is not None:
        st.append(("ALIGN", (alinear_centro_desde, 1), (-1, -1), "CENTER"))
    if zebra:
        for i in range(1, len(datos)):
            if i % 2 == 0:
                st.append(("BACKGROUND", (0, i), (-1, i), GRIS_CL))
    if extra:
        st += extra
    t.setStyle(TableStyle(st))
    return t


def _firmas(meta):
    linea = "_" * 38
    datos = [[linea, linea],
             [f"Elaborado por: {meta['auditor'] or ''}", "Revisado / Visto bueno"],
             ["Responsable de la auditoría", "Jefatura del área"]]
    t = Table(datos, colWidths=[ANCHO / 2] * 2)
    t.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                           ("FONTSIZE", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, 0), 40),
                           ("TEXTCOLOR", (0, 2), (-1, 2), GRIS)]))
    return t


# ---------------------------------------------------------------- informe
def exportar_pdf(aud, tipo, destino, entidad="", auditor="", anios=None, detalle=False, progreso=None):
    res_all = aud.resumen(tipo)
    anios = anios or [r["anio"] for r in res_all]
    res = [r for r in res_all if r["anio"] in anios]
    meta = dict(entidad=entidad.strip(), auditor=auditor.strip(), tipo=tipo,
                archivo=os.path.basename(aud.ruta), fecha=fecha_larga())

    doc = BaseDocTemplate(destino, pagesize=PAGE, leftMargin=MARGEN, rightMargin=MARGEN,
                          topMargin=16 * mm, bottomMargin=14 * mm,
                          title=f"Informe de auditoría de correlativos – {tipo}",
                          author=meta["auditor"] or meta["entidad"], subject="Auditoría de correlativos")
    frame = Frame(MARGEN, 14 * mm, ANCHO, PAGE[1] - 30 * mm, id="f", leftPadding=0, rightPadding=0,
                  topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame])])
    s = []

    tot_reg = sum(r["registros"] for r in res)
    tot_dist = sum(r["distintos"] for r in res)
    tot_ult = sum(r["ultimo"] for r in res)
    tot_falt = sum(r["faltantes"] for r in res)
    tot_rep = sum(r["repetidos"] for r in res)
    revisar = [r for r in aud.por_revisar(tipo) if r.y in anios or r.y is None]
    comp_gl = tot_dist / tot_ult if tot_ult else None
    periodo = f"{min(anios)} – {max(anios)}" if anios else "—"

    # ---------- portada
    s.append(Spacer(1, 38 * mm))
    if meta["entidad"]:
        s.append(Paragraph(esc(meta["entidad"]).upper(), ParagraphStyle(
            "ent", fontName="Helvetica-Bold", fontSize=12, textColor=GRIS, leading=15)))
        s.append(Spacer(1, 6 * mm))
    s.append(Paragraph("Informe de auditoría de correlativos", E["titulo"]))
    s.append(Spacer(1, 3 * mm))
    s.append(Paragraph(f"{esc(tipo)} · período {periodo}", E["subtit"]))
    s.append(Spacer(1, 14 * mm))
    port = [["Archivo auditado", meta["archivo"]], ["Fecha de emisión", meta["fecha"]],
            ["Años auditados", f"{len(res)} ({periodo})"], ["Registros evaluados", miles(tot_reg)]]
    if meta["auditor"]:
        port.append(["Responsable", meta["auditor"]])
    tp = Table(port, colWidths=[55 * mm, 120 * mm])
    tp.setStyle(TableStyle([("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                            ("FONTSIZE", (0, 0), (-1, -1), 10), ("TEXTCOLOR", (0, 0), (0, -1), AZUL),
                            ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
                            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
    s.append(tp)
    s.append(PageBreak())

    # ---------- 1. resumen ejecutivo
    s.append(Paragraph("1. Resumen ejecutivo", E["h1"]))
    s.append(_kpis([("Registros evaluados", miles(tot_reg)), ("Números distintos", miles(tot_dist)),
                    ("Números faltantes", miles(tot_falt)), ("Registros repetidos", miles(tot_rep)),
                    ("Por revisar", miles(len(revisar))), ("Completitud global", pct(comp_gl))]))
    s.append(Spacer(1, 4 * mm))
    peores = sorted([r for r in res if r["faltantes"]], key=lambda r: -r["faltantes"])[:3]
    alto_r = [str(r["anio"]) for r in res if riesgo(r["completitud"])[0] == "ALTO"]
    completos = [str(r["anio"]) for r in res if r["faltantes"] == 0]
    txt = (f"Se revisó la numeración correlativa de <b>{miles(tot_reg)}</b> registros de «{esc(tipo)}» "
           f"correspondientes a <b>{len(res)}</b> años ({periodo}). Considerando que cada año debe contener una "
           f"secuencia continua desde el N° 1 hasta el último número emitido, se identificaron "
           f"<b>{miles(tot_falt)}</b> números faltantes, lo que representa una completitud global de "
           f"<b>{pct(comp_gl)}</b>.")
    if peores:
        txt += " Los años con mayor cantidad de faltantes son " + ", ".join(
            f"{r['anio']} ({miles(r['faltantes'])})" for r in peores) + "."
    if alto_r:
        txt += f" Presentan nivel de riesgo <b>alto</b> (completitud menor a 90%): {', '.join(alto_r)}."
    if completos:
        txt += f" Años con secuencia completa: {', '.join(completos)}."
    if tot_rep:
        txt += (f" Además, {miles(tot_rep)} registros comparten número y año con otro registro, por lo que "
                "corresponde verificar si son duplicados de carga o documentos distintos con numeración repetida.")
    s.append(Paragraph(txt, E["txt"]))
    s.append(Spacer(1, 4 * mm))
    etiq = [str(r["anio"]) for r in res]
    w2 = ANCHO / 2 - 3 * mm
    graf = Table([[_barras([r["faltantes"] for r in res], etiq, w2, 88 * mm, "Números faltantes por año"),
                   _barras([round((r["completitud"] or 0) * 100, 1) for r in res], etiq, w2, 88 * mm,
                           "Completitud por año (%)", color=AZUL, formato="%.0f", max_y=100)]],
                 colWidths=[ANCHO / 2] * 2)
    graf.setStyle(TableStyle([("BOX", (0, 0), (0, 0), 0.4, colors.HexColor("#BFBFBF")),
                              ("BOX", (1, 0), (1, 0), 0.4, colors.HexColor("#BFBFBF")),
                              ("LEFTPADDING", (0, 0), (-1, -1), 4)]))
    s.append(graf)
    s.append(PageBreak())

    # ---------- 2. cuadro consolidado
    s.append(Paragraph("2. Cuadro consolidado por año", E["h1"]))
    filas = [["Año", "Registros", "N° distintos", "Último N°", "Faltantes", "Registros\nrepetidos",
              "Sin número", "Completitud", "Nivel de riesgo", "Situación"]]
    extra = []
    for i, r in enumerate(res, start=1):
        niv, col, colcl = riesgo(r["completitud"])
        filas.append([r["anio"], miles(r["registros"]), miles(r["distintos"]), miles(r["ultimo"]),
                      miles(r["faltantes"]), miles(r["repetidos"]), r["sin_numero"], pct(r["completitud"]),
                      niv, "Completo" if r["faltantes"] == 0 else "Con faltantes"])
        extra += [("BACKGROUND", (8, i), (8, i), colcl), ("TEXTCOLOR", (8, i), (8, i), col),
                  ("FONTNAME", (8, i), (8, i), "Helvetica-Bold")]
        if r["faltantes"]:
            extra.append(("TEXTCOLOR", (4, i), (4, i), ROJO))
    filas.append(["TOTAL", miles(tot_reg), miles(tot_dist), miles(tot_ult), miles(tot_falt), miles(tot_rep),
                  sum(r["sin_numero"] for r in res), pct(comp_gl), "", ""])
    n = len(filas) - 1
    extra += [("FONTNAME", (0, n), (-1, n), "Helvetica-Bold"), ("BACKGROUND", (0, n), (-1, n), AZUL_CL)]
    s.append(_tabla(filas, [22 * mm, 26 * mm, 26 * mm, 24 * mm, 24 * mm, 26 * mm, 22 * mm, 26 * mm, 30 * mm,
                            ANCHO - 226 * mm], extra=extra))
    s.append(Spacer(1, 3 * mm))
    s.append(Paragraph("Nivel de riesgo según completitud: <b>bajo</b> desde 98%; <b>medio</b> de 90% a menos de 98%; "
                       "<b>alto</b> menor a 90%. Completitud = números distintos registrados ÷ último número del año.",
                       E["peq"]))
    s.append(PageBreak())

    # ---------- 3. detalle por año
    for k, r in enumerate(res):
        y = r["anio"]
        seq = aud.secuencia(tipo, y)
        falt = [f[0] for f in seq if f[2] == "FALTA"]
        rep = [f for f in seq if f[2] == "REPETIDO"]
        unicos = sum(1 for f in seq if f[2] == "CONFIRMADO")
        niv, col, colcl = riesgo(r["completitud"])

        cab = Table([[Paragraph(f"3.{k + 1}  Resultados del año {y}", E["h1"]),
                      Paragraph(f"<font color='{col.hexval().replace('0x', '#')}'><b>Riesgo {niv}</b></font>",
                                ParagraphStyle("rg", fontName="Helvetica", fontSize=11, alignment=TA_RIGHT))]],
                    colWidths=[ANCHO * 0.6, ANCHO * 0.4])
        cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                 ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("LINEBELOW", (0, 0), (-1, 0), 2, col)]))
        s.append(cab)
        s.append(Spacer(1, 3 * mm))
        s.append(_kpis([("Registros", miles(r["registros"])), ("Último N°", miles(r["ultimo"])),
                        ("N° distintos", miles(r["distintos"])), ("Faltantes", miles(r["faltantes"])),
                        ("Registros repetidos", miles(r["repetidos"])), ("Completitud", pct(r["completitud"]))]))
        s.append(Spacer(1, 3 * mm))

        vals, cats = _tramos(seq, r["ultimo"])
        g = Table([[_torta(unicos, len(rep), len(falt), ANCHO * 0.40 - 6, 56 * mm),
                    _barras(vals, cats, ANCHO * 0.60 - 6, 56 * mm, "Faltantes por tramo de numeración",
                            rot=35)]], colWidths=[ANCHO * 0.40, ANCHO * 0.60])
        g.setStyle(TableStyle([("BOX", (0, 0), (0, 0), 0.4, colors.HexColor("#BFBFBF")),
                               ("BOX", (1, 0), (1, 0), 0.4, colors.HexColor("#BFBFBF")),
                               ("LEFTPADDING", (0, 0), (-1, -1), 3), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        s.append(g)

        # hallazgo
        s.append(Paragraph("Hallazgo", E["h2"]))
        if r["ultimo"] == 0:
            h = "No se registraron documentos con número para este año."
        elif not falt:
            h = (f"La secuencia del N° 1 al N° {r['ultimo']} está completa: todos los números cuentan con al menos "
                 "un registro.")
        else:
            h = (f"De la secuencia esperada del N° 1 al N° {r['ultimo']} no se encontraron <b>{len(falt)}</b> "
                 f"números ({pct(len(falt) / r['ultimo'])} del total esperado).")
            if vals:
                imax = max(range(len(vals)), key=lambda i: vals[i])
                if vals[imax] >= max(3, len(falt) * 0.3):
                    h += f" La mayor concentración de faltantes está en el tramo {cats[imax]} ({vals[imax]})."
            nums = sorted(f[0] for f in seq if isinstance(f[0], int) and f[1] > 0)
            if len(nums) >= 2:
                saltos = [(a, b) for a, b in zip(nums, nums[1:]) if b - a > 40]
                if saltos:
                    a, b = max(saltos, key=lambda t: t[1] - t[0])
                    h += (f" Se observa un salto de numeración del N° {a} al N° {b}; verificar si los números "
                          "posteriores corresponden a errores de digitación.")
        if rep:
            h += (f" Hay {len(rep)} números registrados más de una vez ({r['repetidos']} registros en total).")
        if r["sin_numero"]:
            h += f" {r['sin_numero']} registro(s) no tienen número de documento."
        s.append(Paragraph(h, E["txt"]))

        s.append(Paragraph(f"Números faltantes ({len(falt)})", E["h2"]))
        s.append(Paragraph(rangos(falt) or "Ninguno.", E["num"]))
        if rep:
            s.append(Paragraph(f"Números repetidos ({len(rep)})", E["h2"]))
            s.append(Paragraph(rangos([f[0] for f in rep]), E["num"]))

        obs = [x for x in revisar if x.y == y]
        if obs:
            tit_obs = Paragraph(f"Observaciones de registro ({len(obs)})", E["h2"])
            filas = [["id", "Título registrado", "N°", "Año", "Estado", "Observación", "Nota del auditor"]]
            for x in obs:
                det = x.problema if x.problema and x.problema != x.estado else x.estado
                filas.append([x.id, Paragraph(esc(x.titulo), E["celda"]), x.n if x.n is not None else "—", x.y,
                              x.estado, Paragraph(esc(DESCRIPCION.get(det, "")), E["celda"]),
                              Paragraph(esc(x.nota or ""), E["celda"])])
            t_obs = _tabla(filas, [22 * mm, 55 * mm, 14 * mm, 14 * mm, 28 * mm, 85 * mm, ANCHO - 218 * mm],
                           alinear_centro_desde=None, extra=[("ALIGN", (2, 1), (4, -1), "CENTER")])
            if len(obs) <= 15:
                s.append(KeepTogether([tit_obs, t_obs]))
            else:
                s.append(CondPageBreak(40 * mm))
                s.append(tit_obs)
                s.append(t_obs)

        if detalle and seq:
            s.append(PageBreak())
            s.append(Paragraph(f"Año {y} – detalle de la secuencia", E["h2"]))
            bloques = 3
            por_bloque = int(math.ceil(len(seq) / bloques))
            cab_d = []
            for _ in range(bloques):
                cab_d += ["N°", "Estado", "Título"]
            filas = [cab_d]
            extra = []
            for i in range(por_bloque):
                fila = []
                for b in range(bloques):
                    j = b * por_bloque + i
                    if j < len(seq):
                        n_, v_, est_, tit_, _ = seq[j]
                        fila += [n_, est_, tit_.split(" | ")[0][:34]]
                        c0 = b * 3
                        if est_ == "FALTA":
                            extra += [("BACKGROUND", (c0, i + 1), (c0 + 2, i + 1), ROJO_CL),
                                      ("TEXTCOLOR", (c0, i + 1), (c0 + 1, i + 1), ROJO)]
                        elif est_ == "REPETIDO":
                            extra.append(("BACKGROUND", (c0, i + 1), (c0 + 2, i + 1), AMBAR_CL))
                    else:
                        fila += ["", "", ""]
                filas.append(fila)
            wcol = ANCHO / bloques
            t = _tabla(filas, [12 * mm, 20 * mm, wcol - 32 * mm] * bloques, zebra=False,
                       alinear_centro_desde=None, extra=extra + [("FONTSIZE", (0, 0), (-1, -1), 6.5),
                                                                 ("TOPPADDING", (0, 0), (-1, -1), 1),
                                                                 ("BOTTOMPADDING", (0, 0), (-1, -1), 1)])
            s.append(t)
        s.append(PageBreak())
        if progreso:
            progreso((k + 1) / len(res))

    # ---------- 4. conclusiones y recomendaciones
    s.append(Paragraph("4. Conclusiones y recomendaciones", E["h1"]))
    concl = [
        f"La numeración de «{esc(tipo)}» presenta una completitud global de {pct(comp_gl)}, con "
        f"{miles(tot_falt)} números faltantes en {sum(1 for r in res if r['faltantes'])} de {len(res)} años auditados.",
    ]
    if alto_r:
        concl.append(f"Los años {', '.join(alto_r)} tienen nivel de riesgo alto y requieren atención prioritaria.")
    if tot_rep:
        concl.append(f"Existen {miles(tot_rep)} registros con número repetido dentro del mismo año.")
    if revisar:
        concl.append(f"{len(revisar)} registros presentan títulos incompletos o con formato irregular.")
    recom = [
        "Ubicar en el archivo físico o en el sistema de trámite los documentos correspondientes a los números "
        "faltantes y registrarlos; si un número fue anulado o no se emitió, dejar constancia documentada.",
        "Verificar los registros con número repetido para determinar si se trata de cargas duplicadas "
        "(depurarlas) o de documentos distintos con la misma numeración (regularizar).",
        "Corregir los títulos incompletos o con formato irregular, usando el formato estándar "
        "NNN-AAAA-SIGLA.",
        "Revisar los saltos grandes de numeración, que suelen indicar errores de digitación del número.",
        "Repetir esta auditoría periódicamente (por ejemplo, cada trimestre) para detectar faltantes a tiempo.",
    ]
    s.append(Paragraph("Conclusiones", E["h2"]))
    for i, c in enumerate(concl, 1):
        s.append(Paragraph(f"{i}. {c}", E["txt"]))
        s.append(Spacer(1, 1.5 * mm))
    s.append(Paragraph("Recomendaciones", E["h2"]))
    for i, c in enumerate(recom, 1):
        s.append(Paragraph(f"{i}. {c}", E["txt"]))
        s.append(Spacer(1, 1.5 * mm))
    s.append(Paragraph("Criterios aplicados", E["h2"]))
    s.append(Paragraph(
        "El número y el año se obtienen del título de cada registro. Cada tipo de documento y año se audita como una "
        "secuencia independiente que debe iniciar en el N° 1 y terminar en el mayor número registrado. Un número es "
        "<b>faltante</b> si ningún registro lo contiene; es <b>repetido</b> si más de un registro tiene el mismo "
        "número y año. Cuando el título no tiene año o lo tiene incompleto, el sistema lo deduce de los registros "
        "vecinos y lo marca para confirmación; las correcciones manuales del auditor se indican como «corregido».",
        E["txt"]))
    s.append(KeepTogether([Spacer(1, 6 * mm), _firmas(meta)]))

    doc.build(s, canvasmaker=hacer_canvas(meta))
    return destino

