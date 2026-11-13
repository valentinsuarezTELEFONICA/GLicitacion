# -*- coding: utf-8 -*-
"""Lectura de fuentes (PLACSP, BOE, TED, feeds extra) y descarga de pliegos."""
from __future__ import annotations

import html as _html
import re
import shutil
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path

import requests

import configuracion as C
import ubicacion
from analisis import Documento, Licitacion, Lote
from textos import norm

NS_ATOM = "{http://www.w3.org/2005/Atom}"
SESION = requests.Session()
SESION.headers["User-Agent"] = "Mozilla/5.0 (compatible; licitaciones-diarias/2.0; uso interno)"


# ------------------------------------------------------------------ utilidades
def get(url: str, **kw) -> requests.Response:
    kw.setdefault("timeout", 60)
    ultimo = None
    for i in range(3):
        try:
            r = SESION.get(url, **kw)
            if r.status_code in (429, 500, 502, 503, 504):
                raise requests.HTTPError(f"HTTP {r.status_code}")
            return r
        except requests.RequestException as e:
            ultimo = e
            time.sleep(2 * (i + 1))
    raise ultimo


def parse_fecha(s: str | None) -> datetime | None:
    if not s:
        return None
    s = str(s).strip().replace("Z", "+00:00")
    if len(s) > 10 and s[10] not in "T ":      # p. ej. "2026-11-10+01:00" (TED): solo la fecha
        s = s[:10]
    try:
        return datetime.fromisoformat(s).replace(tzinfo=None)
    except ValueError:
        pass
    try:
        return datetime.fromisoformat(s[:10])
    except ValueError:
        return None


def _num(s) -> float | None:
    try:
        return float(s) if s not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _ln(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _todos(el, nombre: str):
    return [e for e in el.iter() if _ln(e.tag) == nombre] if el is not None else []


def _primero(el, nombre: str, defecto: str = "") -> str:
    if el is None:
        return defecto
    for e in el.iter():
        if _ln(e.tag) == nombre and (e.text or "").strip():
            return e.text.strip()
    return defecto


def _primer_elemento(el, nombre: str):
    if el is None:
        return None
    for e in el.iter():
        if _ln(e.tag) == nombre:
            return e
    return None


def _hijos(el, nombre: str):
    return [c for c in el if _ln(c.tag) == nombre] if el is not None else []


def _hijo(el, nombre: str):
    h = _hijos(el, nombre)
    return h[0] if h else None


# ------------------------------------------------------------------ PLACSP (Atom + CODICE)
TIPOS_CONTRATO = {"1": "Suministros", "2": "Servicios", "3": "Obras", "21": "Servicios",
                  "22": "Servicios especiales", "31": "Concesión de obras", "32": "Concesión de servicios",
                  "40": "Colaboración público-privada", "7": "Administrativo especial", "8": "Privado"}
PROCEDIMIENTOS = {"1": "Abierto", "2": "Restringido", "3": "Negociado sin publicidad",
                  "4": "Negociado con publicidad", "5": "Diálogo competitivo", "6": "Contrato menor",
                  "7": "Derivado de acuerdo marco", "8": "Concurso de proyectos", "9": "Abierto simplificado",
                  "10": "Asociación para la innovación", "11": "Derivado de sistema dinámico",
                  "12": "Licitación con negociación", "100": "Normas internas", "999": "Otros"}
UNIDADES = {"DAY": "días", "MON": "meses", "ANN": "años", "WEE": "semanas", "HUR": "horas"}


def _importes(budget) -> tuple[float | None, float | None, float | None]:
    """(sin IVA, con IVA, valor estimado) de un BudgetAmount."""
    return (_num(_primero(budget, "TaxExclusiveAmount")), _num(_primero(budget, "TotalAmount")),
            _num(_primero(budget, "EstimatedOverallContractAmount")))


def _cpvs(el) -> list:
    out, vistos = [], set()
    for x in _todos(el, "ItemClassificationCode"):
        c = (x.text or "").strip()
        if c and c not in vistos:
            vistos.add(c)
            out.append(c)
    return out


def _ubicacion_placsp(cfs, pp, organo: str, titulo: str) -> dict:
    loc = _primer_elemento(pp, "RealizedLocation")
    party = _primer_elemento(cfs, "LocatedContractingParty")
    dir_ = _primer_elemento(party, "PostalAddress")
    nuts = _primero(loc, "CountrySubentityCode")
    prov = _primero(loc, "CountrySubentity")
    ciudad = _primero(loc, "CityName") or _primero(dir_, "CityName")
    pais = _primero(loc, "IdentificationCode") or _primero(dir_, "IdentificationCode")
    if not (nuts or prov):
        nuts = _primero(dir_, "CountrySubentityCode")
        prov = _primero(dir_, "CountrySubentity")
    return ubicacion.resolver(nuts=nuts, provincia=prov, ciudad=ciudad, pais=pais, texto=f"{organo} {titulo}")


def entrada_placsp(e) -> Licitacion | None:
    cfs = next((x for x in e if _ln(x.tag) == "ContractFolderStatus"), None)
    if cfs is None:
        return None
    url = ""
    for l in e.findall(NS_ATOM + "link"):
        if l.get("rel") in (None, "alternate"):
            url = l.get("href", "")

    pp = _hijo(cfs, "ProcurementProject")
    if pp is None:
        pp = _primer_elemento(cfs, "ProcurementProject")
    sin_iva, con_iva, valor = _importes(_primer_elemento(pp, "BudgetAmount"))
    notas = []
    presupuesto = sin_iva
    if presupuesto is None and con_iva is not None:
        presupuesto = con_iva
        notas.append("El anuncio solo indica el importe con IVA; se muestra ese.")

    lotes = []
    for n, lp in enumerate(_hijos(cfs, "ProcurementProjectLot"), 1):
        pj = _hijo(lp, "ProcurementProject")
        if pj is None:
            continue
        l_sin, l_con, _ = _importes(_primer_elemento(pj, "BudgetAmount"))
        lotes.append(Lote(numero=_primero(lp, "ID") or str(n), nombre=_primero(pj, "Name"),
                          descripcion=_primero(pj, "Description"), cpv=_cpvs(pj),
                          presupuesto=l_sin if l_sin is not None else l_con))

    plazo = _primer_elemento(cfs, "TenderSubmissionDeadlinePeriod")
    fecha_limite = None
    if plazo is not None:
        d = _primero(plazo, "EndDate")
        h = _primero(plazo, "EndTime")[:8]
        fecha_limite = parse_fecha(f"{d}T{h}" if h else d)

    docs = []
    for nombre, tipo in (("LegalDocumentReference", "PCAP"), ("TechnicalDocumentReference", "PPT"),
                         ("AdditionalDocumentReference", "Anexo")):
        for ref in _todos(cfs, nombre):
            uri = _primero(ref, "URI")
            if uri:
                docs.append(Documento(tipo, _primero(ref, "FileName") or _primero(ref, "ID") or tipo, uri))

    dm = _primer_elemento(_primer_elemento(pp, "PlannedPeriod"), "DurationMeasure")
    duracion = ""
    if dm is not None and (dm.text or "").strip():
        duracion = f"{dm.text.strip()} {UNIDADES.get(dm.get('unitCode', ''), dm.get('unitCode', ''))}".strip()

    codigo_tipo = _primero(pp, "TypeCode")
    titulo = _primero(pp, "Name") or (e.findtext(NS_ATOM + "title") or "").strip()
    organo = _primero(_primer_elemento(cfs, "LocatedContractingParty"), "Name")
    ub = _ubicacion_placsp(cfs, pp, organo, titulo)
    actualizada = parse_fecha(e.findtext(NS_ATOM + "updated"))
    return Licitacion(
        id=(e.findtext(NS_ATOM + "id") or "").strip(), fuente="PLACSP", titulo=titulo,
        expediente=_primero(cfs, "ContractFolderID"), organo=organo,
        descripcion=_primero(pp, "Description"), tipo=TIPOS_CONTRATO.get(codigo_tipo, codigo_tipo),
        procedimiento=PROCEDIMIENTOS.get(_primero(_primer_elemento(cfs, "TenderingProcess"), "ProcedureCode"), ""),
        duracion=duracion, presupuesto=presupuesto, valor_estimado=valor, fecha_limite=fecha_limite,
        fecha_publicacion=parse_fecha(_primero(cfs, "IssueDate")) or actualizada, actualizada=actualizada,
        cpv=_cpvs(cfs), estado=_primero(cfs, "ContractFolderStatusCode"), url=url,
        pais=ub["pais"], ccaa=ub["ccaa"], provincia=ub["provincia"], ciudad=ub["ciudad"],
        ubicacion_origen=ub["origen"], lotes=lotes, documentos=docs, notas=notas,
    )


def leer_placsp(desde: datetime, log=print) -> list[Licitacion]:
    url, out, paginas, vistas = C.PLACSP_FEED, [], 0, set()
    while url and paginas < C.PLACSP_MAX_PAGINAS and url not in vistas:
        vistas.add(url)
        r = get(url)
        r.raise_for_status()
        raiz = ET.fromstring(r.content)
        paginas += 1
        siguiente = next((l.get("href") for l in raiz.findall(NS_ATOM + "link") if l.get("rel") == "next"), None)
        mas_antigua = None
        for e in raiz.findall(NS_ATOM + "entry"):
            lic = entrada_placsp(e)
            if not lic or not lic.actualizada:
                continue
            if mas_antigua is None or lic.actualizada < mas_antigua:
                mas_antigua = lic.actualizada
            if lic.actualizada >= desde:
                out.append(lic)
        log(f"    página {paginas}: {len(out)} anuncios en el periodo")
        if mas_antigua and mas_antigua < desde:
            break
        url = siguiente
    return out


# ------------------------------------------------------------------ BOE (API de datos abiertos)
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
         "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}


def _items_boe(nodo, seccion="", depto=""):
    if isinstance(nodo, dict):
        if "departamento" in nodo and "codigo" in nodo:
            seccion = str(nodo.get("codigo"))
        if "epigrafe" in nodo and "nombre" in nodo:
            depto = str(nodo.get("nombre"))
        if "identificador" in nodo and "titulo" in nodo:
            yield seccion, depto, nodo
        for v in nodo.values():
            yield from _items_boe(v, seccion, depto)
    elif isinstance(nodo, list):
        for v in nodo:
            yield from _items_boe(v, seccion, depto)


def _url_de(valor) -> str:
    if isinstance(valor, dict):
        return str(valor.get("texto") or valor.get("url") or "")
    return str(valor or "")


def parsear_boe(datos: dict) -> list[Licitacion]:
    out = []
    for seccion, depto, it in _items_boe(datos):
        if not seccion.startswith("5"):
            continue
        titulo = str(it["titulo"]).strip()
        nt = norm(titulo)
        if "licitaci" not in nt or "formalizaci" in nt or "adjudicaci" in nt:
            continue
        ident = str(it["identificador"])
        pdf, html = _url_de(it.get("url_pdf")), _url_de(it.get("url_html"))
        organo = depto
        m = re.match(r"anuncio de licitaci[oó]n de:\s*(.+?)\.\s*objeto:", titulo, re.I)
        if m:
            organo = m.group(1)
        exp = re.search(r"expediente:\s*(.+?)\.?\s*$", titulo, re.I)
        obj = re.search(r"objeto:\s*(.+?)(?:\.\s*expediente:.*)?\s*$", titulo, re.I | re.S)
        objeto = obj.group(1).strip().rstrip(".") if obj else titulo      # solo el objeto, sin «Anuncio de licitación de…»
        ub = ubicacion.resolver(texto=f"{organo} {objeto}")
        out.append(Licitacion(
            id=ident, fuente="BOE", titulo=objeto, expediente=exp.group(1).strip() if exp else "",
            organo=organo, url=html or pdf, estado="PUB",
            pais=ub["pais"] or "España", ccaa=ub["ccaa"], provincia=ub["provincia"], ubicacion_origen=ub["origen"],
            documentos=[Documento("Anuncio BOE", f"{ident}.pdf", pdf)] if pdf else [],
        ))
    return out


def leer_boe(desde: datetime, log=print) -> list[Licitacion]:
    out, d = [], desde.date()
    while d <= date.today():
        r = get(C.BOE_API.format(fecha=d.strftime("%Y%m%d")), headers={"Accept": "application/json"})
        if r.status_code == 404:
            log(f"    {d:%d/%m}: sin BOE (domingo o aún no publicado)")
        else:
            r.raise_for_status()
            nuevos = parsear_boe(r.json())
            for n in nuevos:
                n.actualizada = datetime(d.year, d.month, d.day)
                n.fecha_publicacion = n.actualizada
            log(f"    {d:%d/%m}: {len(nuevos)} anuncios de licitación")
            out += nuevos
        d += timedelta(days=1)
    return out


def _fecha_es(texto: str) -> datetime | None:
    t = norm(texto)
    m = re.search(r"(\d{1,2})\s+de\s+(" + "|".join(MESES) + r")\s+de\s+(\d{4})", t)
    if m:
        return datetime(int(m.group(3)), MESES[m.group(2)], int(m.group(1)))
    m = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", t)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def _importe_es(texto: str) -> float | None:
    m = re.search(r"(\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d+(?:,\d{1,2})?)\s*(?:euros|eur\b|€)", texto, re.I)
    if not m:
        return None
    try:
        return float(m.group(1).replace(".", "").replace(",", "."))
    except ValueError:
        return None


def extraer_datos_boe(texto: str) -> dict:
    """Plazo, importes y lugar a partir del texto del anuncio. Búsqueda aproximada: hay que verificarla."""
    t = norm(texto)
    out = {}
    m = re.search(r"(plazo de presentacion|fecha limite de presentacion|presentacion de ofertas|"
                  r"presentacion de proposiciones|recepcion de ofertas)", t)
    if m:
        ventana = texto[m.end(): m.end() + 350]
        f = _fecha_es(ventana)
        if f:
            h = re.search(r"(\d{1,2})[:.](\d{2})\s*(?:horas|h\b)", ventana)
            if h and int(h.group(1)) < 24:
                f = f.replace(hour=int(h.group(1)), minute=int(h.group(2)))
            out["fecha_limite"] = f
    m = re.search(r"valor estimado", t)
    if m:
        out["valor_estimado"] = _importe_es(texto[m.end(): m.end() + 200])
    m = re.search(r"presupuesto base de licitacion", t)
    if m:
        out["presupuesto"] = _importe_es(texto[m.end(): m.end() + 250])
    m = re.search(r"lugar de (ejecucion|entrega|prestacion)[^:]{0,40}:", t)
    if m:
        out["lugar"] = texto[m.end(): m.end() + 160]
    return {k: v for k, v in out.items() if v}


def enriquecer_boe(lic: Licitacion) -> None:
    """Abre el anuncio del BOE y rellena plazo, importes y lugar si los encuentra."""
    if lic.fuente != "BOE" or not lic.url or not lic.url.startswith("http") or lic.url.lower().endswith(".pdf"):
        return
    r = get(lic.url)
    if r.status_code != 200:
        return
    bruto = r.content.decode("utf-8", errors="replace")
    bruto = re.sub(r"(?is)<(script|style).*?</\1>", " ", bruto)
    texto = re.sub(r"\s+", " ", _html.unescape(re.sub(r"<[^>]+>", " ", bruto)))
    datos = extraer_datos_boe(texto)
    if not datos:
        return
    lic.fecha_limite = lic.fecha_limite or datos.get("fecha_limite")
    lic.presupuesto = lic.presupuesto or datos.get("presupuesto")
    lic.valor_estimado = lic.valor_estimado or datos.get("valor_estimado")
    lic.importe_interes = lic.importe_interes or lic.presupuesto
    if datos.get("lugar") and not lic.ccaa:
        ub = ubicacion.resolver(texto=datos["lugar"])
        if ub["ccaa"]:
            lic.ccaa, lic.provincia, lic.pais, lic.ubicacion_origen = ub["ccaa"], ub["provincia"], ub["pais"], "inferido"
    lic.notas.append("Plazo, importes y lugar leídos del texto del anuncio del BOE: verifícalos en el anuncio.")


# ------------------------------------------------------------------ TED (API v3)
def _texto_ted(v) -> str:
    if isinstance(v, str):
        return v
    if isinstance(v, dict):
        for k in ("spa", "SPA", "es", "eng", "ENG"):
            if k in v:
                return _texto_ted(v[k])
        return _texto_ted(next(iter(v.values()), ""))
    if isinstance(v, list):
        return _texto_ted(v[0]) if v else ""
    return "" if v is None else str(v)


def _lista_ted(v) -> list:
    if v is None:
        return []
    return v if isinstance(v, list) else [v]


def parsear_ted(datos: dict) -> list[Licitacion]:
    out = []
    for n in datos.get("notices", []):
        num = str(n.get("publication-number", ""))
        if not num:
            continue
        tipo_aviso = norm(_texto_ted(n.get("notice-type")))
        if tipo_aviso.startswith(("can", "pin")):       # adjudicaciones y avisos previos
            continue
        links = n.get("links") or {}
        html = _texto_ted(links.get("html")) or _texto_ted(links.get("htmlDirect")) \
            or f"https://ted.europa.eu/es/notice/-/detail/{num}"
        pdf = _texto_ted(links.get("pdf")) or _texto_ted(links.get("pdfs"))
        cpv = n.get("classification-cpv") or []
        if isinstance(cpv, str):
            cpv = [cpv]
        titulo = _texto_ted(n.get("notice-title")) or num
        organo = _texto_ted(n.get("buyer-name"))

        valores = [_num(_texto_ted(x)) for x in _lista_ted(n.get("estimated-value-lot"))]
        titulos = [_texto_ted(x) for x in _lista_ted(n.get("title-lot"))]
        descs = [_texto_ted(x) for x in _lista_ted(n.get("description-lot"))]
        lotes = []
        if len(titulos) > 1:
            for i, t in enumerate(titulos):
                lotes.append(Lote(numero=str(i + 1), nombre=t, descripcion=descs[i] if i < len(descs) else "",
                                  presupuesto=valores[i] if len(valores) == len(titulos) else None))
        total = sum(v for v in valores if v) or None

        pais = _texto_ted(n.get("place-of-performance-country-lot")) or _texto_ted(n.get("buyer-country"))
        ciudad = _texto_ted(n.get("place-of-performance-city-lot")) or _texto_ted(n.get("buyer-city"))
        ub = ubicacion.resolver(pais=pais, ciudad=ciudad, texto=f"{organo} {titulo}")
        plazo = n.get("deadline-receipt-tender-date-lot")
        fecha_pub = parse_fecha(_texto_ted(n.get("publication-date")))
        out.append(Licitacion(
            id=num, fuente="TED", titulo=titulo, expediente=num, organo=organo, url=html, estado="PUB",
            cpv=[str(c) for c in cpv], fecha_limite=parse_fecha(_texto_ted(plazo)),
            fecha_publicacion=fecha_pub, actualizada=fecha_pub,
            tipo=_texto_ted(n.get("contract-nature")), procedimiento=_texto_ted(n.get("procedure-type")),
            valor_estimado=total, lotes=lotes, pais=ub["pais"], ccaa=ub["ccaa"], provincia=ub["provincia"],
            ciudad=ub["ciudad"], ubicacion_origen=ub["origen"],
            documentos=[Documento("Anuncio TED", f"{num}.pdf", pdf)] if pdf else [],
            notas=["Importe de TED: valor estimado (suma de lotes)."] if total else [],
        ))
    return out


def _ted_post(campos: list, pagina: int, consulta: str, limite: int = 100):
    return SESION.post(C.TED_API, json={"query": consulta, "fields": campos, "page": pagina, "limit": limite},
                       timeout=60)


def _ted_campos_validos(consulta: str, log) -> list:
    """Averigua qué campos opcionales acepta TED y descarta los que no reconoce."""
    base = list(C.TED_CAMPOS_BASE)
    r = _ted_post(base, 1, consulta, 1)
    if r.status_code != 200:
        raise RuntimeError(f"TED respondió {r.status_code}: {r.text[:200]}")
    ok = list(base)
    for campo in C.TED_CAMPOS_EXTRA:
        if _ted_post(ok + [campo], 1, consulta, 1).status_code == 200:
            ok.append(campo)
        else:
            log(f"    (TED no reconoce el campo «{campo}»: se ignora)")
    return ok


def leer_ted(desde: datetime, log=print, meta: dict | None = None) -> list[Licitacion]:
    meta = meta if meta is not None else {}
    consulta = f"buyer-country=ESP AND publication-date>={desde:%Y%m%d}"
    campos = meta.get("ted_campos")
    if not campos:
        r = _ted_post(list(C.TED_CAMPOS_BASE) + list(C.TED_CAMPOS_EXTRA), 1, consulta, 1)
        campos = (list(C.TED_CAMPOS_BASE) + list(C.TED_CAMPOS_EXTRA)) if r.status_code == 200 \
            else _ted_campos_validos(consulta, log)
        meta["ted_campos"] = campos
    out = []
    for pagina in range(1, C.TED_MAX_PAGINAS + 1):
        r = _ted_post(campos, pagina, consulta)
        if r.status_code == 400 and pagina == 1 and meta.get("ted_campos"):
            meta.pop("ted_campos", None)             # la lista recordada ya no vale: se recalcula mañana
        r.raise_for_status()
        lote = parsear_ted(r.json())
        out += lote
        log(f"    página {pagina}: {len(lote)} anuncios")
        if len(lote) < 100:
            break
    return out


# ------------------------------------------------------------------ feeds RSS/Atom genéricos
def leer_feed_extra(nombre: str, url: str, desde: datetime, log=print) -> list[Licitacion]:
    r = get(url)
    r.raise_for_status()
    raiz = ET.fromstring(r.content)
    out = []
    for it in list(raiz.iter("item")) + list(raiz.iter(NS_ATOM + "entry")):
        def t(*nombres):
            for n in nombres:
                v = it.findtext(n)
                if v:
                    return v.strip()
            return ""
        enlace = t("link")
        if not enlace:
            l = it.find(NS_ATOM + "link")
            enlace = l.get("href", "") if l is not None else ""
        fecha = None
        crudo = t("pubDate", NS_ATOM + "updated", NS_ATOM + "published")
        if crudo:
            try:
                from email.utils import parsedate_to_datetime
                fecha = parsedate_to_datetime(crudo).replace(tzinfo=None)
            except (TypeError, ValueError):
                fecha = parse_fecha(crudo)
        if fecha and fecha < desde:
            continue
        titulo = t("title", NS_ATOM + "title")
        desc = re.sub(r"<[^>]+>", " ", t("description", NS_ATOM + "summary", NS_ATOM + "content"))
        ub = ubicacion.resolver(texto=f"{titulo} {desc}")
        out.append(Licitacion(
            id=t("guid", "id", NS_ATOM + "id") or enlace, fuente=nombre, titulo=titulo, descripcion=desc,
            url=enlace, actualizada=fecha, fecha_publicacion=fecha, estado="PUB",
            pais=ub["pais"], ccaa=ub["ccaa"], provincia=ub["provincia"], ubicacion_origen=ub["origen"],
            documentos=[Documento("Anuncio", "enlace", enlace)] if enlace else [],
        ))
    log(f"    {len(out)} anuncios")
    return out


# ------------------------------------------------------------------ descarga de pliegos
def _nombre_seguro(s: str, defecto: str) -> str:
    s = re.sub(r"[^\w.\- ]+", "_", urllib.parse.unquote(s)).strip(" ._")[:80] or defecto
    return s if s.lower().endswith(".pdf") else s + ".pdf"


class _Reintentable(Exception):
    """Fallo pasajero del servidor (error 5xx, 429, página de error en vez del PDF): merece otro intento."""


def _con_paciencia(fn, veces: int = 6, pausa: float = 0.5):
    """Windows no deja abrir ni borrar un archivo mientras el antivirus o OneDrive lo están revisando
    (WinError 32): se espera un poco y se vuelve a probar."""
    for n in range(veces):
        try:
            return fn()
        except PermissionError:
            if n == veces - 1:
                raise
            time.sleep(pausa * (n + 1))


def _borrar(ruta: Path) -> None:
    try:
        _con_paciencia(lambda: ruta.unlink(missing_ok=True))
    except OSError:
        pass                                   # si no se puede borrar, se queda; no es grave


def _cabecera(ruta: Path) -> bytes:
    def leer():
        with open(ruta, "rb") as f:            # se cierra antes de devolver: así luego se puede borrar
            return f.read(512)
    return _con_paciencia(leer)


def _que_es(cab: bytes, tipo_http: str) -> tuple[str, bool]:
    """Describe lo recibido cuando no es un PDF. Devuelve (texto, si merece reintentar)."""
    t = cab.lstrip()[:200].lower()
    if cab[:2] == b"PK":
        return "es un ZIP o un Word/Excel, no un PDF: ábrelo desde «Original»", False
    if cab[:4] == b"\xd0\xcf\x11\xe0":
        return "es un documento de Office antiguo, no un PDF: ábrelo desde «Original»", False
    if t.startswith((b"<!doctype", b"<html", b"<?xml")) or b"<html" in t or "html" in tipo_http:
        return "el servidor devolvió una página web en lugar del PDF", True
    if t.startswith(b"error"):
        return f"el servidor devolvió un error ({cab[:60].decode('latin-1', 'replace').strip()})", True
    return f"no es un PDF (tipo {tipo_http or 'desconocido'})", False


def descargar(doc: Documento, carpeta: Path, intentos: int = 3) -> None:
    """Descarga un pliego. Si el servidor falla de forma pasajera, lo reintenta con pausas crecientes.
    Deja el motivo real del fallo en doc.error."""
    carpeta.mkdir(parents=True, exist_ok=True)
    base = doc.nombre if doc.nombre and doc.nombre != "enlace" else Path(urllib.parse.urlparse(doc.url).path).name
    destino = carpeta / _nombre_seguro(f"{doc.tipo}_{base}", "documento.pdf")
    motivo = ""
    for intento in range(1, intentos + 1):
        tipo_http = ""
        try:
            if doc.url.startswith("file://"):
                shutil.copy(urllib.request.url2pathname(urllib.parse.urlparse(doc.url).path), destino)
            else:
                with SESION.get(doc.url, stream=True, timeout=90) as r:
                    if r.status_code in (429, 500, 502, 503, 504):
                        raise _Reintentable(f"el servidor respondió con el error {r.status_code}")
                    r.raise_for_status()
                    tipo_http = (r.headers.get("Content-Type") or "").lower()
                    limite = C.MAX_MB_POR_DOCUMENTO * 1024 * 1024
                    tam = 0
                    with open(destino, "wb") as f:
                        for trozo in r.iter_content(65536):
                            tam += len(trozo)
                            if tam > limite:
                                raise ValueError(f"supera {C.MAX_MB_POR_DOCUMENTO} MB")
                            f.write(trozo)
            cab = _cabecera(destino)
            if cab[:5] == b"%PDF-":
                doc.ruta, doc.error = str(destino), ""
                return
            _borrar(destino)
            motivo, reintentar = _que_es(cab, tipo_http)
            if not reintentar:
                doc.error = motivo
                return
        except _Reintentable as e:
            _borrar(destino)
            motivo = str(e)
        except (requests.ConnectionError, requests.Timeout) as e:
            _borrar(destino)
            motivo = f"sin respuesta del servidor ({type(e).__name__})"
        except Exception as e:  # noqa: BLE001 — un documento fallido no debe parar el informe
            _borrar(destino)
            doc.error = f"no se pudo descargar ({type(e).__name__}: {e})"
            return
        if intento < intentos:
            time.sleep(4 * intento)            # 4 s, 8 s...: da tiempo al servidor a recuperarse
    doc.error = f"no se pudo descargar tras {intentos} intentos: {motivo}"
