# -*- coding: utf-8 -*-
"""
Genera la web estática para los comerciales dentro de la carpeta de informes (la que se comparte en OneDrive):

    informes/
        index.html            <- la página (se copia desde la carpeta «web» del programa)
        datos/indice.js       <- lista ligera de todas las licitaciones (la carga la página al abrir)
        datos/f/<id>.js       <- ficha completa de cada licitación (se carga al abrir su ficha)
        config/*.js           <- panel de configuración (lee y guarda palabras_clave.xlsx); se copia de web/config

Se usan ficheros .js y no .json porque los navegadores no dejan leer .json de disco al abrir una página
con doble clic (file://), pero sí cargar scripts. Así la web funciona sin servidor y sin internet.
Los ficheros que no han cambiado no se reescriben, para que OneDrive no sincronice cambios inútiles.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path

DIA = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def slug(clave: str) -> str:
    return hashlib.sha1(clave.encode("utf-8")).hexdigest()[:12]


def _js(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def escribir_si_cambia(ruta: Path, texto: str) -> bool:
    """Escritura atómica (fichero temporal + renombrado) y solo si el contenido es distinto."""
    ruta = Path(ruta)
    try:
        if ruta.exists() and ruta.read_text(encoding="utf-8") == texto:
            return False
    except OSError:
        pass
    ruta.parent.mkdir(parents=True, exist_ok=True)
    tmp = ruta.with_name(ruta.name + ".tmp")
    tmp.write_text(texto, encoding="utf-8")
    os.replace(tmp, ruta)
    return True


def copiar_si_cambia(origen: Path, destino: Path) -> bool:
    """Copia exacta (byte a byte, sin tocar los finales de línea), atómica y solo si cambia."""
    contenido = Path(origen).read_bytes()
    destino = Path(destino)
    try:
        if destino.exists() and destino.read_bytes() == contenido:
            return False
    except OSError:
        pass
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_name(destino.name + ".tmp")
    tmp.write_bytes(contenido)
    os.replace(tmp, destino)
    return True


def _indice(f: dict) -> dict:
    return {
        "id": slug(f["clave"]), "titulo": f.get("titulo", ""), "organo": f.get("organo", ""),
        "ccaa": f.get("ccaa", ""), "provincia": f.get("provincia", ""), "ciudad": f.get("ciudad", ""),
        "pais": f.get("pais", ""), "aprox": f.get("ubicacion_origen") == "inferido",
        "encaje": f.get("encaje", ""), "puntos": f.get("puntos", 0), "novedad": f.get("novedad", ""),
        "cambios": f.get("cambios", ""), "fuente": f.get("fuente", ""),
        "presupuesto": f.get("presupuesto"), "interes": f.get("importe_interes"),
        "lotes": f.get("n_lotes", 0), "lotesInteres": f.get("n_lotes_interes"),
        "limite": f.get("fecha_limite", ""), "dia": f.get("dia", ""), "primera": f.get("primera", ""),
        "palabras": sorted({p["termino"] for p in f.get("palabras", [])}),
        "tipo": f.get("tipo", ""), "expediente": f.get("expediente", ""),
        "tambien": [d["fuente"] for d in f.get("duplicada_en", [])],
    }


def _dias(raiz: Path, fichas: list) -> list:
    """Un registro por día con lo recogido (nuevas) o actualizado, y los enlaces a su Excel y PDF si existen."""
    dias: dict = {}
    for f in fichas:
        if f.get("primera"):
            dias.setdefault(f["primera"], {"nuevas": 0, "actualizadas": 0})["nuevas"] += 1
        if f.get("dia") and f.get("dia") != f.get("primera"):
            dias.setdefault(f["dia"], {"nuevas": 0, "actualizadas": 0})["actualizadas"] += 1
    if raiz.exists():
        for c in raiz.iterdir():
            if c.is_dir() and DIA.match(c.name):
                dias.setdefault(c.name, {"nuevas": 0, "actualizadas": 0})
    out = []
    for d, n in sorted(dias.items(), reverse=True):
        x = raiz / d / f"licitaciones_{d}.xlsx"
        p = raiz / d / f"licitaciones_{d}.pdf"
        out.append({"dia": d, **n, "excel": f"{d}/{x.name}" if x.exists() else "",
                    "pdf": f"{d}/{p.name}" if p.exists() else ""})
    return out


IMAGENES = (".png", ".jpg", ".jpeg", ".svg", ".webp", ".gif")


def exportar(raiz: Path, fichas: list, info: dict, programa: Path, empresa: str = "", log=print,
             logos: dict | None = None, empresa_defecto: str = "") -> Path:
    """Escribe datos/, config/ e index.html dentro de «raiz». «info» describe la última ejecución.
    «logos» = {"claro": imagen, "oscuro": imagen}, nombres de archivos de web/config (vacío = icono por defecto)."""
    carpeta_config = Path(programa) / "web" / "config"
    imagenes = sorted(p.name for p in carpeta_config.glob("*") if p.suffix.lower() in IMAGENES) \
        if carpeta_config.exists() else []
    raiz = Path(raiz)
    datos = raiz / "datos"
    carpeta_f = datos / "f"
    carpeta_f.mkdir(parents=True, exist_ok=True)

    esperados, escritos = set(), 0
    for f in fichas:
        s = slug(f["clave"])
        esperados.add(s + ".js")
        cuerpo = ("window.__FICHAS=window.__FICHAS||{};window.__FICHAS[" + _js(s) + "]=" + _js(f) + ";")
        escritos += escribir_si_cambia(carpeta_f / (s + ".js"), cuerpo)
    for viejo in carpeta_f.glob("*.js"):                    # fichas que ya no existen
        if viejo.name not in esperados:
            viejo.unlink(missing_ok=True)

    hist = raiz / "HISTORICO_licitaciones.xlsx"
    resumen = {
        "generado": datetime.now().isoformat(timespec="minutes"),
        "empresa": empresa,
        "empresaPorDefecto": empresa_defecto or empresa,   # el de configuracion.py, si se deja vacío en el panel
        "logo": {k: v for k, v in (logos or {}).items() if v in imagenes},
        "logos": imagenes,                     # imágenes disponibles, para elegirlas en Configuración
        "ultimoDia": info.get("dia") or max((f.get("dia", "") for f in fichas), default=""),
        "hora": info.get("hora", ""),
        "fuentes": [{"nombre": n, "ok": bool(ok), "texto": t} for n, ok, t in info.get("fuentes", [])],
        "descartadas": info.get("descartadas", 0), "conocidas": info.get("conocidas", 0),
        "historico": hist.name if hist.exists() else "",
        "dias": _dias(raiz, fichas),
        "fichas": [_indice(f) for f in sorted(fichas, key=lambda f: (-f.get("puntos", 0), f.get("titulo", "")))],
    }
    escribir_si_cambia(datos / "indice.js", "window.DATOS=" + _js(resumen) + ";")

    plantilla = Path(programa) / "web" / "index.html"
    if not plantilla.exists():
        raise FileNotFoundError(f"Falta la plantilla de la web: {plantilla}")
    copiar_si_cambia(plantilla, raiz / "index.html")
    for archivo in carpeta_config.glob("*") if carpeta_config.exists() else []:
        # panel de configuración (lee y guarda el Excel) y logos disponibles
        if archivo.is_file() and (archivo.suffix.lower() == ".js" or archivo.suffix.lower() in IMAGENES):
            copiar_si_cambia(archivo, raiz / "config" / archivo.name)
    log(f"  Web actualizada: {len(fichas)} licitaciones ({escritos} fichas escritas o cambiadas) en {raiz / 'index.html'}")
    return raiz / "index.html"
