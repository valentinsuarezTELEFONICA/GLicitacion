# -*- coding: utf-8 -*-
"""
Una «ficha» es la licitación ya analizada, en forma de diccionario simple (se guarda en registro.json).
Los informes (PDF y Excel) y, más adelante, el motor web, se generan a partir de las fichas.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import configuracion as C
from analisis import Licitacion, etiqueta_encaje

FECHA = "%d/%m/%Y"


def iso(d: datetime | None) -> str:
    return d.isoformat(timespec="minutes") if d else ""


def fecha(s: str) -> datetime | None:
    try:
        return datetime.fromisoformat(s) if s else None
    except ValueError:
        return None


def rel(ruta: str, raiz: Path) -> str:
    """Ruta de un fichero relativa a la carpeta de informes (con «/»)."""
    if not ruta:
        return ""
    try:
        return Path(ruta).resolve().relative_to(Path(raiz).resolve()).as_posix()
    except ValueError:
        return Path(ruta).as_posix()


def filas_requisitos(lic: Licitacion, raiz: Path) -> list:
    filas = []
    for concepto in C.CONCEPTOS:
        for a in lic.analisis:
            hits = a["hits"].get(concepto)
            if not hits:
                continue
            filas.append({
                "concepto": concepto, "tipo": a["tipo"], "documento": a["nombre"],
                "ruta_rel": rel(a.get("ruta", ""), raiz), "url": a.get("url", ""),
                "paginas": [n for n, _ in hits[:C.PAGINAS_POR_CONCEPTO]],
                "mas": max(0, len(hits) - C.PAGINAS_POR_CONCEPTO),
                "extracto": hits[0][1], "pagina_extracto": hits[0][0],
            })
    return filas


def ficha_de(lic: Licitacion, dia: str, primera: str, raiz: Path, clave: str, marca: str) -> dict:
    return {
        "clave": clave, "marca": marca, "id": lic.id, "fuente": lic.fuente, "dia": dia, "primera": primera,
        "novedad": lic.novedad, "cambios": lic.cambios, "puntos": lic.puntos,
        "encaje": etiqueta_encaje(lic.puntos),
        "titulo": lic.titulo, "expediente": lic.expediente, "organo": lic.organo,
        "descripcion": (lic.descripcion or "")[:600], "tipo": lic.tipo, "procedimiento": lic.procedimiento,
        "duracion": lic.duracion, "estado": lic.estado,
        "pais": lic.pais, "ccaa": lic.ccaa, "provincia": lic.provincia, "ciudad": lic.ciudad,
        "ubicacion_origen": lic.ubicacion_origen,
        "presupuesto": lic.presupuesto, "valor_estimado": lic.valor_estimado,
        "n_lotes": len(lic.lotes), "n_lotes_interes": lic.n_lotes_interes, "importe_interes": lic.importe_interes,
        "fecha_publicacion": iso(lic.fecha_publicacion), "fecha_limite": iso(lic.fecha_limite),
        "actualizada": iso(lic.actualizada), "cpv": list(lic.cpv), "palabras": lic.palabras,
        "motivos": lic.motivos, "url": lic.url,
        "documentos": [{"tipo": d.tipo, "nombre": d.nombre, "url": d.url,
                        "ruta_rel": rel(d.ruta, raiz), "error": d.error} for d in lic.documentos],
        "lotes": [{"numero": l.numero, "nombre": l.nombre, "presupuesto": l.presupuesto, "cpv": l.cpv,
                   "interes": l.interes, "puntos": l.puntos, "palabras": l.palabras} for l in lic.lotes],
        "requisitos": filas_requisitos(lic, raiz), "ia": lic.ia, "notas": list(lic.notas),
        "duplicada_en": list(lic.duplicada_en),
    }


def comparar(previa: dict, nueva: dict) -> str:
    """Texto con lo que ha cambiado entre dos fichas (vacío si nada relevante)."""
    c = []

    def f(s):
        d = fecha(s)
        return d.strftime(FECHA + (" %H:%M" if d.hour or d.minute else "")) if d else "sin fecha"

    def e(v):
        return f"{v:,.0f} €".replace(",", ".") if v is not None else "sin importe"

    if previa.get("fecha_limite") != nueva.get("fecha_limite"):
        c.append(f"Fecha límite: {f(previa.get('fecha_limite'))} -> {f(nueva.get('fecha_limite'))}")
    if previa.get("presupuesto") != nueva.get("presupuesto"):
        c.append(f"Presupuesto: {e(previa.get('presupuesto'))} -> {e(nueva.get('presupuesto'))}")
    if previa.get("valor_estimado") != nueva.get("valor_estimado"):
        c.append(f"Valor estimado: {e(previa.get('valor_estimado'))} -> {e(nueva.get('valor_estimado'))}")
    if previa.get("estado") != nueva.get("estado"):
        c.append(f"Estado: {previa.get('estado')} -> {nueva.get('estado')}")
    if previa.get("titulo") != nueva.get("titulo"):
        c.append("Cambia el título")
    antes = {d["url"] for d in previa.get("documentos", [])}
    nuevos = [d["nombre"] for d in nueva.get("documentos", []) if d["url"] not in antes]
    if nuevos:
        c.append("Documentos nuevos: " + ", ".join(nuevos[:3]))
    return "; ".join(c)
