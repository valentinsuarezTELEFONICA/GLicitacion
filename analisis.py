# -*- coding: utf-8 -*-
"""Modelo de datos, puntuación de relevancia (con lotes) y análisis de pliegos (con páginas)."""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import configuracion as C
import palabras
from textos import norm  # noqa: F401  (re-exportado: otros módulos lo importan de aquí)


@dataclass
class Documento:
    tipo: str            # PCAP, PPT, Anexo, Anuncio...
    nombre: str
    url: str
    ruta: str = ""       # ruta local si se descargó
    error: str = ""


@dataclass
class Lote:
    numero: str
    nombre: str = ""
    descripcion: str = ""
    presupuesto: float | None = None     # sin IVA
    cpv: list = field(default_factory=list)
    interes: bool = False
    puntos: int = 0
    palabras: list = field(default_factory=list)


@dataclass
class Licitacion:
    id: str
    fuente: str
    titulo: str
    expediente: str = ""
    organo: str = ""
    descripcion: str = ""
    tipo: str = ""
    procedimiento: str = ""
    duracion: str = ""
    presupuesto: float | None = None        # presupuesto base sin IVA
    valor_estimado: float | None = None
    fecha_limite: datetime | None = None
    fecha_publicacion: datetime | None = None
    actualizada: datetime | None = None
    cpv: list = field(default_factory=list)
    estado: str = ""
    url: str = ""
    pais: str = ""
    ccaa: str = ""
    provincia: str = ""
    ciudad: str = ""
    ubicacion_origen: str = ""
    lotes: list = field(default_factory=list)
    documentos: list = field(default_factory=list)
    novedad: str = "Nueva"
    cambios: str = ""
    puntos: int = 0
    motivos: list = field(default_factory=list)
    palabras: list = field(default_factory=list)      # {"grupo","termino","donde":[...]}
    importe_interes: float | None = None
    n_lotes_interes: int | None = None
    analisis: list = field(default_factory=list)
    ia: dict | None = None
    notas: list = field(default_factory=list)
    duplicada_en: list = field(default_factory=list)  # {"fuente","url"}


# ------------------------------------------------------------------ puntuación
def etiqueta_encaje(puntos: int, reglas=None) -> str:
    r = reglas or palabras.ACTUAL
    if puntos >= r.alto:
        return "ALTO"
    if puntos >= r.minima:
        return "MEDIO"
    return "BAJO"


def _zonas(lic: Licitacion) -> dict:
    z = {"título": norm(lic.titulo), "descripción": norm(lic.descripcion)}
    for lote in lic.lotes:
        z[f"lote {lote.numero}"] = norm(f"{lote.nombre} {lote.descripcion}")
    return z


def _buscar(reglas, zonas: dict) -> list:
    out = []
    for t in reglas.terminos:
        donde = [z for z, txt in zonas.items() if txt and t.regex.search(txt)]
        if donde:
            out.append({"grupo": t.grupo, "termino": t.texto, "donde": donde})
    return out


def _puntos(reglas, hallazgos: list, cpvs: list) -> tuple[int, list]:
    motivos, grupos = [], {}
    for h in hallazgos:
        grupos.setdefault(h["grupo"], []).append(h["termino"])
    suma = sum(reglas.grupos.get(g, 0) for g in grupos)
    for g, terminos in grupos.items():
        motivos.append(f"{g}: {', '.join(terminos[:4])}")
    puntos = min(suma, reglas.tope)
    mejor = (0, "", "")
    for codigo in cpvs:
        for prefijo, (peso, etiqueta) in reglas.cpv.items():
            if str(codigo).startswith(prefijo) and peso > mejor[0]:
                mejor = (peso, etiqueta, str(codigo))
    if mejor[0]:
        puntos += mejor[0]
        motivos.append(f"CPV {mejor[2]} ({mejor[1]})")
    return puntos, motivos


def puntuar(lic: Licitacion, reglas=None) -> None:
    """Puntúa la licitación y cada lote; calcula qué parte del importe corresponde a lo que interesa."""
    r = reglas or palabras.ACTUAL
    zonas = _zonas(lic)
    lic.palabras = _buscar(r, zonas)
    puntos, motivos = _puntos(r, lic.palabras, lic.cpv)

    texto = f"{zonas['título']} {zonas['descripción']}"
    excl = []
    for t in r.exclusiones:
        if t.regex.search(texto):
            excl.append(t.texto)
    if excl:
        puntos -= r.penalizacion * min(len(excl), 2)
        motivos.append("Posible descarte: " + ", ".join(excl[:3]))

    if lic.presupuesto is not None:
        if r.importe_min and lic.presupuesto < r.importe_min:
            puntos -= 30
            motivos.append("Importe por debajo de tu mínimo")
        if r.importe_max and lic.presupuesto > r.importe_max:
            puntos -= 30
            motivos.append("Importe por encima de tu máximo")
    lic.puntos = max(0, min(100, puntos))
    lic.motivos = motivos

    # ---- lotes: cuáles interesan y cuánto suman
    for lote in lic.lotes:
        zl = {f"lote {lote.numero}": norm(f"{lote.nombre} {lote.descripcion}")}
        lote.palabras = _buscar(r, zl)
        lote.puntos, _ = _puntos(r, lote.palabras, lote.cpv)
        if any(t.regex.search(txt) for t in r.exclusiones for txt in zl.values()):
            lote.puntos = max(0, lote.puntos - r.penalizacion)      # p. ej. «obra de reforma de la sala»
        lote.interes = lote.puntos >= r.minima_lote
    if lic.lotes:
        interes = [l for l in lic.lotes if l.interes]
        lic.n_lotes_interes = len(interes)
        if interes:
            importes = [l.presupuesto for l in interes if l.presupuesto is not None]
            lic.importe_interes = sum(importes) if importes else None
            if len(importes) < len(interes):
                lic.notas.append("Algún lote de interés no indica importe.")
        else:
            lic.importe_interes = None
            lic.notas.append("Ningún lote destaca por su texto o CPV: revisa a mano qué lotes te interesan.")
    else:
        lic.n_lotes_interes = None
        lic.importe_interes = lic.presupuesto   # un solo lote: todo el contrato


# ------------------------------------------------------------------ análisis de pliegos
def _extracto(orig: str, ini: int, fin: int, antes: int = 110, despues: int = 260) -> str:
    a, b = max(0, ini - antes), min(len(orig), fin + despues)
    s = re.sub(r"\s+", " ", orig[a:b]).strip()
    return ("…" if a > 0 else "") + s + ("…" if b < len(orig) else "")


def analizar_pdf(ruta: str) -> dict:
    """Lee el PDF página a página y localiza en qué páginas aparece cada concepto."""
    try:
        import pymupdf as fitz  # PyMuPDF (nombre actual)
    except ImportError:
        import fitz  # PyMuPDF (versiones antiguas)

    paginas: dict[int, str] = {}
    with fitz.open(ruta) as doc:
        for i, pagina in enumerate(doc, 1):
            paginas[i] = pagina.get_text("text") or ""

    util = sum(len(t.strip()) for t in paginas.values()) > 200
    res = {"paginas": len(paginas), "hits": {}, "texto_util": util, "textos": paginas}
    if not util:
        return res

    for concepto, patrones in C.CONCEPTOS.items():
        encontrados = []
        for n, t in paginas.items():
            nt = norm(t)
            for patron in patrones:
                m = re.search(patron, nt)
                if m:
                    encontrados.append((n, _extracto(t, m.start(), m.end())))
                    break
        if encontrados:
            res["hits"][concepto] = encontrados
    return res


# ------------------------------------------------------------------ resumen con IA (opcional)
def ia_disponible() -> bool:
    return C.USAR_IA and bool(os.environ.get("ANTHROPIC_API_KEY"))


def resumen_ia(lic: Licitacion) -> dict | None:
    """Pide a Claude un resumen de lo que se pide, con páginas. Devuelve None si no se puede."""
    import requests

    bloques, total = [], 0
    for a in lic.analisis:
        if not a.get("hits"):
            continue
        claves = sorted({n for hs in a["hits"].values() for n, _ in hs[:3]})
        for n in claves:
            t = a["textos"][n].strip()[:3500]
            bloque = f"[DOCUMENTO: {a['tipo']} «{a['nombre']}» — PÁGINA {n}]\n{t}\n"
            if total + len(bloque) > 60000:
                break
            bloques.append(bloque)
            total += len(bloque)
    if not bloques:
        return None

    importe = f"{lic.presupuesto:,.0f} € (sin IVA)" if lic.presupuesto else "no indicado"
    plazo = lic.fecha_limite.strftime("%d/%m/%Y %H:%M") if lic.fecha_limite else "no indicado"
    prompt = f"""Eres analista de licitaciones públicas en España. Perfil de la empresa:
{C.PERFIL_EMPRESA}

Licitación: {lic.titulo}
Órgano: {lic.organo} · Presupuesto base: {importe} · Fecha límite: {plazo}

Extractos del pliego (con número de página):
{''.join(bloques)}

Responde SOLO con un JSON válido, en español, con esta forma:
{{"encaje": "alto|medio|bajo", "motivo": "1-2 frases sobre por qué encaja o no con la empresa",
 "resumen": "máximo 120 palabras: qué se pide realmente",
 "requisitos": [{{"concepto": "...", "detalle": "...", "documento": "PCAP|PPT|...", "pagina": 0}}],
 "riesgos": ["..."], "documentacion": ["certificados, solvencia, etc. que habría que aportar"]}}
Usa únicamente lo que aparece en los extractos; si algo no consta, no lo inventes."""

    r = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": os.environ["ANTHROPIC_API_KEY"],
                 "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={"model": C.MODELO_IA, "max_tokens": 1800,
              "messages": [{"role": "user", "content": prompt}]},
        timeout=120,
    )
    r.raise_for_status()
    texto = r.json()["content"][0]["text"]
    ini, fin = texto.find("{"), texto.rfind("}")
    return json.loads(texto[ini:fin + 1])
