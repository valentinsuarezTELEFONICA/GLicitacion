# -*- coding: utf-8 -*-
"""Genera el Excel diario, el Excel histórico acumulado y el PDF a partir de las fichas."""
from __future__ import annotations

import os
import urllib.parse
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import configuracion as C
from fichas import fecha
from textos import euros

COLOR_ENCAJE = {"ALTO": "#2e7d32", "MEDIO": "#c77700", "BAJO": "#757575"}
RELLENO_ENCAJE = {"ALTO": "C8E6C9", "MEDIO": "FFE0B2", "BAJO": "EEEEEE"}
AZUL = "#1a56b0"
FUENTE = "Arial"


# ------------------------------------------------------------------ utilidades comunes
def limpiar(s) -> str:
    """Deja solo caracteres que la fuente estándar del PDF puede pintar."""
    return str(s or "").encode("cp1252", "replace").decode("cp1252")


def esc(s) -> str:
    return escape(limpiar(s))


def dias_restantes(f: dict, hoy: date):
    d = fecha(f.get("fecha_limite", ""))
    return (d.date() - hoy).days if d else None


def texto_plazo(f: dict, hoy: date) -> str:
    d = fecha(f.get("fecha_limite", ""))
    if not d:
        return "Ver anuncio"
    n = (d.date() - hoy).days
    txt = d.strftime("%d/%m/%Y") + (d.strftime(" %H:%M") if d.hour or d.minute else "")
    if n < 0:
        return f"{txt} (cerrada)"
    return f"{txt} ({'hoy' if n == 0 else 'mañana' if n == 1 else f'{n} días'})"


def ubicacion_txt(f: dict) -> str:
    partes = [p for p in (f.get("ciudad"), f.get("provincia"), f.get("ccaa"), f.get("pais")) if p]
    vistos, out = set(), []
    for p in partes:
        if p.lower() not in vistos:
            vistos.add(p.lower())
            out.append(p)
    txt = ", ".join(out) or "No consta"
    return txt + (" (deducida del texto)" if f.get("ubicacion_origen") == "inferido" and out else "")


def palabras_txt(f: dict) -> str:
    return "; ".join(f"{p['termino']} [{', '.join(p['donde'])}]" for p in f.get("palabras", []))


def requisitos_txt(f: dict) -> str:
    partes = []
    for r in f.get("requisitos", []):
        pags = ",".join(str(n) for n in r["paginas"]) + ("+" if r["mas"] else "")
        partes.append(f"{r['concepto']}: {r['tipo']} p.{pags}")
    return "; ".join(partes)


def lotes_txt(f: dict) -> str:
    if not f.get("n_lotes"):
        return "Sin lotes"
    return f"{f.get('n_lotes_interes') or 0} de {f['n_lotes']}"


def doc_local(f: dict) -> str:
    docs = [d for d in f.get("documentos", []) if d.get("ruta_rel")]
    return docs[0]["ruta_rel"] if docs else ""


def doc_original(f: dict) -> str:
    docs = [d for d in f.get("documentos", []) if d.get("url") and d.get("tipo") in ("PCAP", "PPT", "Anexo")]
    return docs[0]["url"] if docs else ""


def enlace_relativo(raiz: Path, ruta_rel: str, carpeta: Path) -> str:
    """Enlace de Excel a un fichero, relativo a la carpeta del libro (se puede mover la carpeta entera)."""
    destino = os.path.relpath(Path(raiz) / ruta_rel, carpeta).replace("\\", "/")
    return urllib.parse.quote(destino, safe="/")


def guardar_excel(wb: Workbook, ruta: Path, log=print) -> Path:
    try:
        wb.save(ruta)
        return ruta
    except PermissionError:
        alt = ruta.with_name(ruta.stem + "_nuevo" + ruta.suffix)
        wb.save(alt)
        log(f"  ¡Atención! {ruta.name} está abierto en Excel. Se ha guardado como {alt.name}; ciérralo para que se actualice el original.")
        return alt


# ------------------------------------------------------------------ Excel
def _columnas(historico: bool) -> list:
    """(clave, título, ancho, función, formato, tipo_enlace)"""
    def num(k):
        return lambda f: f.get(k)

    cols = [
        ("encaje", "Encaje", 9, lambda f: f["encaje"], None, None),
        ("puntos", "Puntos", 8, num("puntos"), "0", None),
        ("objeto", "Objeto", 50, num("titulo"), None, None),
        ("organo", "Órgano de contratación", 30, num("organo"), None, None),
        ("novedad", "Novedad", 12, num("novedad"), None, None),
        ("cambios", "Qué ha cambiado", 34, num("cambios"), None, None),
        ("palabras", "Palabras clave detectadas [dónde]", 40, palabras_txt, None, None),
        ("ccaa", "Comunidad autónoma", 19, num("ccaa"), None, None),
        ("provincia", "Provincia", 16, num("provincia"), None, None),
        ("ciudad", "Ciudad", 16, num("ciudad"), None, None),
        ("pais", "País", 10, num("pais"), None, None),
        ("total", "Presupuesto total s/IVA (€)", 16, num("presupuesto"), "#,##0", None),
        ("interes", "Importe de lo que te interesa s/IVA (€)", 19, num("importe_interes"), "#,##0", None),
        ("lotes", "Lotes de interés", 11, lotes_txt, None, None),
        ("valor", "Valor estimado (€)", 15, num("valor_estimado"), "#,##0", None),
        ("limite", "Fecha límite de presentación", 17, lambda f: fecha(f.get("fecha_limite", "")), "dd/mm/yyyy hh:mm", None),
        ("dias", "Días restantes", 10, None, "0", None),
        ("publicacion", "Publicación / actualización", 14, lambda f: fecha(f.get("fecha_publicacion", "")), "dd/mm/yyyy", None),
        ("tipo", "Tipo de contrato", 14, num("tipo"), None, None),
        ("procedimiento", "Procedimiento", 18, num("procedimiento"), None, None),
        ("duracion", "Duración", 12, num("duracion"), None, None),
        ("cpv", "CPV", 20, lambda f: ", ".join(f.get("cpv", [])[:6]), None, None),
        ("porque", "Por qué encaja (puntuación)", 45, lambda f: "; ".join(f.get("motivos", [])), None, None),
        ("requisitos", "Requisitos hallados en el pliego (páginas)", 50, requisitos_txt, None, None),
        ("anuncio", "Anuncio", 12, lambda f: "Abrir" if f.get("url") else "", None, "anuncio"),
        ("pliego", "Pliego descargado", 14, lambda f: "Abrir" if doc_local(f) else "", None, "local"),
        ("original", "Pliego original", 14, lambda f: "Abrir" if doc_original(f) else "", None, "original"),
        ("tambien", "También publicada en", 24,
         lambda f: "; ".join(d["fuente"] for d in f.get("duplicada_en", [])), None, None),
        ("expediente", "Expediente", 18, num("expediente"), None, None),
        ("fuente", "Fuente", 10, num("fuente"), None, None),
        ("estado", "Estado", 8, num("estado"), None, None),
        ("dia", "Fecha de búsqueda", 13, lambda f: fecha(f.get("dia", "")), "dd/mm/yyyy", None),
        ("primera", "Primera recogida", 13, lambda f: fecha(f.get("primera", "")), "dd/mm/yyyy", None),
        ("notas", "Avisos", 42,
         lambda f: "; ".join(f.get("notas", []) + [d["error"] for d in f.get("documentos", []) if d.get("error")]), None, None),
    ]
    return cols


def _estilo_cabecera(ws, n):
    azul = PatternFill("solid", fgColor="1A56B0")
    for c in ws[1][:n]:
        c.font = Font(name=FUENTE, bold=True, color="FFFFFF", size=10)
        c.fill = azul
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42


def escribir_oportunidades(ws, fichas: list, raiz: Path, carpeta: Path, historico: bool = False) -> None:
    cols = _columnas(historico)
    ws.append([c[1] for c in cols])
    _estilo_cabecera(ws, len(cols))
    letra = {c[0]: get_column_letter(i) for i, c in enumerate(cols, 1)}
    azul = Font(name=FUENTE, size=10, color="1A56B0", underline="single")
    for f in fichas:
        fila = ws.max_row + 1
        valores = []
        for clave, _, _, fn, _, _ in cols:
            if clave == "dias":
                L = letra["limite"]
                valores.append(f'=IF({L}{fila}="","",INT({L}{fila})-TODAY())')
            else:
                valores.append(fn(f))
        ws.append(valores)
        for i, (clave, _, _, _, fmt, enlace) in enumerate(cols, 1):
            celda = ws.cell(fila, i)
            celda.font = Font(name=FUENTE, size=10)
            celda.alignment = Alignment(vertical="top", wrap_text=True)
            if fmt:
                celda.number_format = fmt
            destino = None
            if enlace == "anuncio" and f.get("url"):
                destino = f["url"]
            elif enlace == "local" and doc_local(f):
                destino = enlace_relativo(raiz, doc_local(f), carpeta)
            elif enlace == "original" and doc_original(f):
                destino = doc_original(f)
            if destino:
                celda.hyperlink = destino
                celda.font = azul
        ws.cell(fila, 1).fill = PatternFill("solid", fgColor=RELLENO_ENCAJE[f["encaje"]])
        ws.cell(fila, 1).font = Font(name=FUENTE, size=10, bold=True)
        if f.get("ubicacion_origen") == "inferido" and f.get("ccaa"):
            ws.cell(fila, list(letra).index("ccaa") + 1).value = f["ccaa"] + " (aprox.)"
    for i, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = c[2]
    ws.freeze_panes = ws.cell(2, list(letra).index("objeto") + 2)
    ws.auto_filter.ref = ws.dimensions
    D = f"{letra['dias']}2:{letra['dias']}{max(ws.max_row, 2)}"
    ws.conditional_formatting.add(D, CellIsRule(operator="between", formula=["0", "5"],
                                                font=Font(name=FUENTE, bold=True, color="B71C1C"),
                                                fill=PatternFill("solid", bgColor="FFCDD2")))
    ws.conditional_formatting.add(D, CellIsRule(operator="lessThan", formula=["0"],
                                                font=Font(name=FUENTE, color="9E9E9E")))


def _hoja_lotes(wb, fichas, raiz, carpeta):
    ws = wb.create_sheet("Lotes")
    ws.append(["Objeto", "Órgano", "Lote", "Descripción del lote", "Presupuesto s/IVA (€)", "CPV",
               "¿De interés?", "Palabras clave [dónde]", "Anuncio"])
    _estilo_cabecera(ws, 9)
    azul = Font(name=FUENTE, size=10, color="1A56B0", underline="single")
    for f in fichas:
        for l in f.get("lotes", []):
            ws.append([f["titulo"][:120], f["organo"], l["numero"], l["nombre"], l["presupuesto"],
                       ", ".join(l["cpv"][:4]), "SÍ" if l["interes"] else "NO",
                       "; ".join(f"{p['termino']}" for p in l["palabras"]), "Abrir" if f.get("url") else ""])
            fila = ws.max_row
            for c in ws[fila]:
                c.font = Font(name=FUENTE, size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)
            ws.cell(fila, 5).number_format = "#,##0"
            ws.cell(fila, 7).fill = PatternFill("solid", fgColor="C8E6C9" if l["interes"] else "EEEEEE")
            if f.get("url"):
                ws.cell(fila, 9).hyperlink = f["url"]
                ws.cell(fila, 9).font = azul
    for i, a in enumerate([45, 28, 8, 55, 16, 22, 11, 40, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _hoja_requisitos(wb, fichas, raiz, carpeta):
    ws = wb.create_sheet("Requisitos y páginas")
    ws.append(["Objeto", "Concepto", "Documento", "Páginas", "Extracto", "Abrir documento"])
    _estilo_cabecera(ws, 6)
    azul = Font(name=FUENTE, size=10, color="1A56B0", underline="single")
    for f in fichas:
        for r in f.get("requisitos", []):
            pags = ", ".join(str(n) for n in r["paginas"]) + (f" (+{r['mas']})" if r["mas"] else "")
            ws.append([f["titulo"][:120], r["concepto"], f"{r['tipo']} · {r['documento']}", pags,
                       f"p.{r['pagina_extracto']}: {r['extracto']}", "Abrir" if r["ruta_rel"] or r["url"] else ""])
            fila = ws.max_row
            for c in ws[fila]:
                c.font = Font(name=FUENTE, size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)
            destino = enlace_relativo(raiz, r["ruta_rel"], carpeta) if r["ruta_rel"] else r["url"]
            if destino:
                ws.cell(fila, 6).hyperlink = destino
                ws.cell(fila, 6).font = azul
    for i, a in enumerate([45, 26, 34, 14, 100, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _hoja_fuentes(wb, ctx):
    ws = wb.create_sheet("Fuentes")
    ws.append(["Fuente", "Resultado"])
    _estilo_cabecera(ws, 2)
    for nombre, ok, texto in ctx["fuentes"]:
        ws.append([nombre, ("OK · " if ok else "ERROR · ") + texto])
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            c.font = Font(name=FUENTE, size=10, color="000000" if c.column == 1 or "OK" in str(fila[1].value)[:3] else "B71C1C")
            c.alignment = Alignment(vertical="top", wrap_text=True)
    if ctx.get("cambios_estado"):
        ws.append([])
        ws.append(["Cambios de estado de licitaciones ya recogidas", ""])
        ws.cell(ws.max_row, 1).font = Font(name=FUENTE, bold=True)
        for c in ctx["cambios_estado"]:
            ws.append([c["titulo"][:100], f"{c['antes']} -> {c['ahora']}"])
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 110


def generar_excel_dia(ruta: Path, fichas: list, ctx: dict, raiz: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Oportunidades"
    escribir_oportunidades(ws, fichas, raiz, ruta.parent)
    _hoja_lotes(wb, fichas, raiz, ruta.parent)
    _hoja_requisitos(wb, fichas, raiz, ruta.parent)
    _hoja_fuentes(wb, ctx)
    return guardar_excel(wb, ruta)


def generar_historico(ruta: Path, fichas: list, raiz: Path) -> Path:
    """Todo lo recogido desde siempre, sin duplicados, más un resumen por día con fórmulas."""
    fichas = sorted(fichas, key=lambda f: (f.get("dia", ""), f.get("puntos", 0)), reverse=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Histórico"
    escribir_oportunidades(ws, fichas, raiz, ruta.parent, historico=True)

    cols = _columnas(True)
    col_primera = get_column_letter([c[0] for c in cols].index("primera") + 1)
    col_encaje = "A"
    r = wb.create_sheet("Resumen por día")
    r.append(["Día", "Oportunidades recogidas por primera vez", "De ellas, encaje ALTO", "De ellas, encaje MEDIO"])
    _estilo_cabecera(r, 4)
    dias = sorted({f["primera"] for f in fichas if f.get("primera")}, reverse=True)
    for d in dias:
        fila = r.max_row + 1
        r.append([fecha(d),
                  f"=COUNTIF('Histórico'!${col_primera}:${col_primera},A{fila})",
                  f"=COUNTIFS('Histórico'!${col_primera}:${col_primera},A{fila},'Histórico'!${col_encaje}:${col_encaje},\"ALTO\")",
                  f"=COUNTIFS('Histórico'!${col_primera}:${col_primera},A{fila},'Histórico'!${col_encaje}:${col_encaje},\"MEDIO\")"])
        r.cell(fila, 1).number_format = "dd/mm/yyyy"
        for c in r[fila]:
            c.font = Font(name=FUENTE, size=10)
    for i, a in enumerate([14, 22, 18, 18], 1):
        r.column_dimensions[get_column_letter(i)].width = a
    r.freeze_panes = "A2"
    return guardar_excel(wb, ruta)


# ------------------------------------------------------------------ PDF
def generar_pdf(ruta: Path, fichas: list, ctx: dict, raiz: Path) -> None:
    hoy = ctx["fecha"]
    base = getSampleStyleSheet()
    s_norm = ParagraphStyle("n", parent=base["Normal"], fontName="Helvetica", fontSize=8.5, leading=11)
    s_peq = ParagraphStyle("p", parent=s_norm, fontSize=7.5, leading=9.5)
    s_h1 = ParagraphStyle("h1", parent=base["Title"], fontName="Helvetica-Bold", fontSize=18, alignment=0)
    s_h2 = ParagraphStyle("h2", parent=s_norm, fontName="Helvetica-Bold", fontSize=11.5, leading=14, spaceBefore=4)
    s_h3 = ParagraphStyle("h3", parent=s_norm, fontName="Helvetica-Bold", fontSize=9, spaceBefore=5, spaceAfter=2)

    def p(t, st=s_norm):
        return Paragraph(t, st)

    def enlace(url, texto):
        if not url:
            return esc(texto)
        return f'<a href="{escape(url, {chr(34): "&quot;"})}" color="{AZUL}">{esc(texto)}</a>'

    def uri(ruta_rel):
        return (Path(raiz) / ruta_rel).resolve().as_uri()

    def pie(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.grey)
        canvas.drawString(15 * mm, 9 * mm, limpiar(f"Informe de licitaciones · {C.EMPRESA} · {hoy:%d/%m/%Y}"))
        canvas.drawRightString(195 * mm, 9 * mm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(str(ruta), pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=15 * mm, bottomMargin=16 * mm,
                            title=f"Licitaciones {hoy:%d/%m/%Y}", author=limpiar(C.EMPRESA))
    h = [p(esc(f"Licitaciones del {hoy:%d/%m/%Y}"), s_h1)]
    if ctx.get("demo"):
        h.append(p('<font color="#b71c1c"><b>DATOS DE DEMOSTRACIÓN: no son licitaciones reales.</b></font>', s_norm))
    nuevas = sum(1 for f in fichas if f["novedad"] == "Nueva")
    h.append(p(esc(f"{C.EMPRESA} · {len(fichas)} oportunidades hoy ({nuevas} nuevas, {len(fichas) - nuevas} con cambios o ya recogidas) · "
                   f"{ctx['descartadas']} anuncios leídos descartados por baja relevancia · {ctx.get('repetidas', 0)} ya conocidas"), s_norm))
    h.append(Spacer(1, 3 * mm))

    filas = [[p("<b>Fuente</b>", s_peq), p("<b>Resultado</b>", s_peq)]]
    for nombre, ok, texto in ctx["fuentes"]:
        color = "#2e7d32" if ok else "#b71c1c"
        filas.append([p(esc(nombre), s_peq), p(f'<font color="{color}">{esc(texto)}</font>', s_peq)])
    t = Table(filas, colWidths=[55 * mm, 125 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                           ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    h += [t, Spacer(1, 5 * mm), p("Resumen", s_h2)]

    cab = ["#", "Encaje", "Órgano y lugar", "Objeto", "Total / de interés", "Plazo"]
    filas = [[p(f"<b>{c}</b>", s_peq) for c in cab]]
    for i, f in enumerate(fichas, 1):
        d = dias_restantes(f, hoy)
        plazo = esc(texto_plazo(f, hoy))
        if d is not None and 0 <= d <= 5:
            plazo = f'<font color="#b71c1c"><b>{plazo}</b></font>'
        lugar = ", ".join(x for x in (f.get("provincia"), f.get("ccaa")) if x)
        interes = "" if f.get("importe_interes") in (None, f.get("presupuesto")) else f"<br/>interés: {esc(euros(f['importe_interes']))}"
        filas.append([
            p(f'<a href="#lic{i}" color="{AZUL}">{i}</a>', s_peq),
            p(f'<font color="{COLOR_ENCAJE[f["encaje"]]}"><b>{f["encaje"]}</b></font><br/>{f["puntos"]} pts', s_peq),
            p(esc(f["organo"][:60]) + (f"<br/><i>{esc(lugar)}</i>" if lugar else ""), s_peq),
            p(esc((f["titulo"][:140] + "…") if len(f["titulo"]) > 140 else f["titulo"]), s_peq),
            p(esc(euros(f.get("presupuesto"))) + interes, s_peq), p(plazo, s_peq)])
    t = Table(filas, colWidths=[7 * mm, 15 * mm, 38 * mm, 62 * mm, 30 * mm, 28 * mm], repeatRows=1)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                           ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    h.append(t)
    if not fichas:
        h.append(p("Hoy no hay nada nuevo que supere el umbral de relevancia.", s_norm))

    for i, f in enumerate(fichas, 1):
        bloque = [Spacer(1, 7 * mm),
                  p(f'<a name="lic{i}"/>{i}. <font color="{COLOR_ENCAJE[f["encaje"]]}">[{f["encaje"]} · {f["puntos"]} pts]</font> '
                    f'{esc(f["titulo"][:300])}', s_h2)]
        meta = [
            ("Órgano", esc(f["organo"])),
            ("Lugar", esc(ubicacion_txt(f))),
            ("Fuente / estado", esc(f"{f['fuente']} · {f['novedad']}" + (f" · exp. {f['expediente']}" if f["expediente"] else ""))),
        ]
        if f.get("cambios"):
            meta.append(("Qué ha cambiado", f'<font color="#b71c1c">{esc(f["cambios"])}</font>'))
        meta += [
            ("Tipo / procedimiento", esc(" · ".join(x for x in (f["tipo"], f["procedimiento"], f["duracion"]) if x) or "—")),
            ("Presupuesto total (s/IVA)", esc(euros(f.get("presupuesto")))),
            ("Lo que te interesa", esc(f"{euros(f.get('importe_interes'))} ({lotes_txt(f)} lotes)"
                                       if f.get("n_lotes") else "Todo el contrato (sin lotes)")),
            ("Valor estimado", esc(euros(f.get("valor_estimado")))),
            ("Fecha límite", esc(texto_plazo(f, hoy))),
            ("CPV", esc(", ".join(f["cpv"][:6]) or "—")),
            ("Palabras clave", esc(palabras_txt(f) or "—")),
            ("Anuncio", enlace(f["url"], "Abrir la licitación en su web") if f.get("url") else "—"),
        ]
        if f.get("duplicada_en"):
            meta.append(("También publicada en", " · ".join(enlace(d["url"], d["fuente"]) for d in f["duplicada_en"])))
        t = Table([[p(f"<b>{k}</b>", s_peq), p(v, s_peq)] for k, v in meta], colWidths=[38 * mm, 142 * mm])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
                               ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                               ("LINEBELOW", (0, 0), (-1, -1), 0.2, colors.lightgrey)]))
        bloque.append(t)

        if f.get("lotes"):
            bloque.append(p("Lotes", s_h3))
            datos = [[p("<b>Lote</b>", s_peq), p("<b>Descripción</b>", s_peq), p("<b>Importe</b>", s_peq), p("<b>¿Interesa?</b>", s_peq)]]
            for l in f["lotes"]:
                datos.append([p(esc(l["numero"]), s_peq), p(esc(l["nombre"][:160]), s_peq), p(esc(euros(l["presupuesto"])), s_peq),
                              p('<font color="#2e7d32"><b>SÍ</b></font>' if l["interes"] else "no", s_peq)])
            t = Table(datos, colWidths=[12 * mm, 108 * mm, 33 * mm, 27 * mm], repeatRows=1)
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                                   ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                                   ("VALIGN", (0, 0), (-1, -1), "TOP")]))
            bloque.append(t)

        if f.get("ia"):
            ia = f["ia"]
            bloque.append(p("Resumen de lo que se pide (IA)", s_h3))
            bloque.append(p(esc(f"Valoración: {ia.get('encaje', '—')}. {ia.get('motivo', '')}")))
            bloque.append(p(esc(ia.get("resumen", ""))))
            if ia.get("riesgos"):
                bloque.append(p("<b>Riesgos:</b> " + esc("; ".join(map(str, ia["riesgos"])))))
            if ia.get("documentacion"):
                bloque.append(p("<b>Documentación a aportar:</b> " + esc("; ".join(map(str, ia["documentacion"])))))

        if f.get("requisitos"):
            bloque.append(p("Qué se pide y dónde está (pulsa la página para abrir el documento)", s_h3))
            datos = [[p("<b>Concepto</b>", s_peq), p("<b>Dónde</b>", s_peq), p("<b>Extracto</b>", s_peq)]]
            for r in f["requisitos"]:
                pags = " ".join(enlace(f"{uri(r['ruta_rel'])}#page={n}" if r["ruta_rel"] else (r["url"] and f"{r['url']}#page={n}"),
                                       f"p.{n}") for n in r["paginas"])
                if r["mas"]:
                    pags += esc(f" (+{r['mas']})")
                datos.append([p(esc(r["concepto"]), s_peq), p(f"<b>{esc(r['tipo'])}</b> {pags}", s_peq),
                              p(esc(f"p.{r['pagina_extracto']}: {r['extracto']}"), s_peq)])
            t = Table(datos, colWidths=[30 * mm, 38 * mm, 112 * mm], repeatRows=1)
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                                   ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                                   ("VALIGN", (0, 0), (-1, -1), "TOP")]))
            bloque.append(t)

        if f.get("documentos"):
            bloque.append(p("Documentos", s_h3))
            for d in f["documentos"]:
                partes = [f"<b>{esc(d['tipo'])}</b> {esc(d['nombre'])}"]
                if d.get("ruta_rel"):
                    partes.append(enlace(uri(d["ruta_rel"]), "copia local"))
                if d.get("url"):
                    partes.append(enlace(d["url"], "original en línea"))
                if d.get("error"):
                    partes.append(f'<font color="#b71c1c">{esc(d["error"])}</font>')
                bloque.append(p(" · ".join(partes), s_peq))
        for n in f.get("notas", []):
            bloque.append(p(f'<font color="#b71c1c">{esc(n)}</font>', s_peq))

        h.append(KeepTogether(bloque[:3]))
        h += bloque[3:]

    doc.build(h, onFirstPage=pie, onLaterPages=pie)
