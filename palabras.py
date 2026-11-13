# -*- coding: utf-8 -*-
"""
Reglas de búsqueda editables: palabras_clave.xlsx

El Excel se crea la primera vez con valores por defecto. Se edita con Excel y se aplica en la
siguiente ejecución. Aquí está también el código que lo lee y lo convierte en reglas.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import configuracion as C
from textos import norm

# ------------------------------------------------------------------ valores por defecto
GRUPOS = [
    ("Realidad virtual, aumentada e inmersiva", 45, "VR, AR, XR, experiencias inmersivas, hologramas"),
    ("Museos, interpretación y turismo", 35, "Museos, centros de interpretación, patrimonio, turismo"),
    ("Audiovisual y pantallas", 30, "Videowalls, pantallas interactivas, proyección, señalización digital"),
    ("Hardware informático", 20, "Ordenadores, servidores, tablets"),
    ("Software y aplicaciones", 25, "Software a medida, apps, páginas y portales web"),
    ("Gestión documental y CMS", 30, "Gestores documentales, gestores de contenidos, digitalización"),
    ("Gamificación y educación digital", 30, "Gamificación, videojuegos, e-learning"),
]

N, E = "Normal", "Exacta"
PALABRAS = [
    ("Realidad virtual, aumentada e inmersiva", t, tipo) for t, tipo in [
        ("realidad virtual", N), ("realidad aumentada", N), ("realidad mixta", N), ("realidad extendida", N),
        ("VR", E), ("XR", E), ("inmersiv", N), ("metaverso", N), ("holograf", N), ("holograma", N),
        ("360 grados", N), ("gemelo digital", N)]
] + [
    ("Museos, interpretación y turismo", t, tipo) for t, tipo in [
        ("museo", N), ("museograf", N), ("centro de interpretación", N), ("centro de visitantes", N),
        ("exposición", N), ("expositiv", N), ("audioguía", N), ("patrimonio cultural", N),
        ("patrimonio histórico", N), ("turístic", N), ("interpretación del patrimonio", N)]
] + [
    ("Audiovisual y pantallas", t, tipo) for t, tipo in [
        ("videowall", N), ("video wall", N), ("muro de video", N), ("pantalla interactiv", N),
        ("pantalla táctil", N), ("pantalla led", N), ("pantalla digital", N), ("audiovisual", N),
        ("videoproyec", N), ("proyección", N), ("mapping", E), ("tótem", N), ("quiosco", N),
        ("señalización digital", N), ("digital signage", N)]
] + [
    ("Hardware informático", t, tipo) for t, tipo in [
        ("equipo informático", N), ("equipamiento informático", N), ("ordenador", N), ("portátil", N),
        ("servidores", E), ("tablet", N), ("workstation", N), ("monitor", N)]
] + [
    ("Software y aplicaciones", t, tipo) for t, tipo in [
        ("software", N), ("desarrollo de aplicaciones", N), ("desarrollo de software", N),
        ("desarrollo de sistemas", N), ("aplicación móvil", N), ("aplicación web", N), ("aplicación informática", N),
        ("app", E), ("apps", E), ("página web", N), ("portal web", N), ("sitio web", N), ("plataforma web", N),
        ("desarrollo web", N)]
] + [
    ("Gestión documental y CMS", t, tipo) for t, tipo in [
        ("gestor documental", N), ("gestión documental", N), ("gestor de contenidos", N), ("CMS", E),
        ("archivo electrónico", N), ("digitalización", N), ("expediente electrónico", N)]
] + [
    ("Gamificación y educación digital", t, tipo) for t, tipo in [
        ("gamific", N), ("serious game", N), ("juego serio", N), ("videojuego", N), ("e-learning", N),
        ("elearning", N), ("aula virtual", N), ("plataforma educativa", N), ("plataforma de formación", N)]
]

CPV = [
    ("48", 30, "Paquetes de software"), ("72", 30, "Servicios informáticos"),
    ("30200", 25, "Equipos informáticos"), ("30213", 25, "Ordenadores"),
    ("30231", 25, "Pantallas y consolas"), ("32320", 30, "Equipo audiovisual"),
    ("32321", 30, "Equipo audiovisual"), ("32322", 30, "Equipo multimedia"),
    ("92520", 35, "Museos"), ("92521", 35, "Servicios de museos"),
    ("92500", 20, "Bibliotecas, archivos y museos"), ("51600", 15, "Instalación de ordenadores"),
]

EXCLUSIONES = [
    ("obra de construcción", N), ("obra de reforma", N), ("obra de rehabilitación", N),
    ("obra de urbanización", N), ("obra civil", N), ("limpieza", N), ("vigilancia y seguridad", N),
    ("mantenimiento de ascensores", N), ("mantenimiento de jardines", N), ("zonas verdes", N),
    ("suministro de gasóleo", N), ("suministro de combustible", N), ("suministro de energía eléctrica", N),
    ("suministro de alimentos", N), ("suministro de medicamentos", N), ("recogida de residuos", N),
    ("transporte escolar", N), ("transporte sanitario", N), ("comedor", N),
]

AJUSTES = [
    ("Puntuación mínima para que aparezca en el informe", 25,
     "Por debajo de esta puntuación la licitación se descarta."),
    ("Puntuación a partir de la cual el encaje es ALTO", 50, "Entre la mínima y esta, el encaje es MEDIO."),
    ("Tope de puntos que pueden dar las palabras clave", 70, "Evita que muchas palabras seguidas lo disparen todo."),
    ("Puntuación mínima para considerar un lote de interés", 20,
     "Se calcula con el texto y el CPV de cada lote por separado."),
    ("Penalización por palabra de exclusión", 40, "Puntos que se restan si aparece una palabra de la hoja Exclusiones."),
    ("Importe mínimo en euros (vacío = sin límite)", None, "Presupuesto base sin IVA. Resta puntos si es menor."),
    ("Importe máximo en euros (vacío = sin límite)", None, "Presupuesto base sin IVA. Resta puntos si es mayor."),
]

# Filas de texto de la hoja Ajustes: se editan en la pestaña «Empresa» del panel de configuración.
# El programa (y el panel) las reconocen por el principio del nombre: «Nombre de la empresa» y «Logo».
AJUSTES_EMPRESA = [
    ("Nombre de la empresa", "", "Aparece en la cabecera de la web y en el PDF. Vacío = el de configuracion.py."),
    ("Logo para el tema claro", "", "Imagen de la carpeta web/config. Vacío = el icono por defecto."),
    ("Logo para el tema oscuro", "", "Imagen de la carpeta web/config. Vacío = el mismo del tema claro."),
]

INSTRUCCIONES = [
    "CÓMO USAR ESTE ARCHIVO",
    "",
    "Aquí decides qué licitaciones te interesan. Edítalo con Excel, guarda y CIERRA el archivo; "
    "se aplica en la siguiente ejecución del programa.",
    "",
    "Hoja «Grupos»: familias de servicio y los puntos que da cada una. Si en una licitación aparece "
    "cualquier palabra de un grupo, suma los puntos del grupo (una sola vez por grupo).",
    "Hoja «Palabras»: las palabras o expresiones que se buscan en el título, la descripción y los lotes. "
    "Puedes añadir filas, cambiarlas o poner «NO» en la columna Activa para desactivarlas sin borrarlas.",
    "   · Tipo «Normal»: acepta terminaciones. «museo» encuentra museo, museos, museográfico; "
    "«pantalla interactiv» encuentra pantallas interactivas. Escribe la raíz de la palabra.",
    "   · Tipo «Exacta»: solo la palabra completa. Úsalo para siglas cortas (VR, XR, CMS, app) que, "
    "si no, aparecerían dentro de otras palabras.",
    "   · No importan mayúsculas ni tildes.",
    "Hoja «CPV»: códigos de clasificación. Escribe el inicio del código (el prefijo «72» cubre todos los "
    "72xxxxxx). Da los puntos indicados; si hay varios, cuenta el mayor.",
    "Hoja «Exclusiones»: palabras que restan puntos (limpieza, obras...). Funcionan como las palabras clave.",
    "Hoja «Ajustes»: umbrales de puntuación e importes mínimo y máximo.",
    "",
    "Cuando cambies palabras y quieras revisar de nuevo lo de los últimos días:",
    "   py licitaciones.py --dias 7 --todo",
    "El informe diario mostrará lo ya recogido como «Ya recogida» y el histórico se actualizará sin duplicados.",
]


# ------------------------------------------------------------------ modelo
@dataclass
class Termino:
    grupo: str
    texto: str
    regex: re.Pattern
    exacta: bool = False


@dataclass
class Reglas:
    grupos: dict = field(default_factory=dict)         # nombre -> puntos
    terminos: list = field(default_factory=list)       # Termino
    cpv: dict = field(default_factory=dict)            # prefijo -> (puntos, etiqueta)
    exclusiones: list = field(default_factory=list)    # Termino
    minima: int = 25
    alto: int = 50
    tope: int = 70
    minima_lote: int = 20
    penalizacion: int = 40
    importe_min: float | None = None
    importe_max: float | None = None
    ruta: Path | None = None
    avisos: list = field(default_factory=list)
    empresa: str = ""            # vacío = se usa EMPRESA de configuracion.py
    logo_claro: str = ""         # nombre de una imagen de web/config; vacío = icono por defecto
    logo_oscuro: str = ""


ACTUAL: Reglas = Reglas()   # reglas en uso; las rellena cargar()


def compilar(texto: str, exacta: bool = False):
    """Convierte una palabra o expresión en una expresión regular tolerante (sin tildes ni mayúsculas)."""
    partes = norm(texto).split()
    if not partes:
        return None
    cuerpo = []
    for i, p in enumerate(partes):
        e = re.escape(p)
        if exacta:
            cuerpo.append(e)
        else:
            cuerpo.append(e + (r"[a-z]*" if i == len(partes) - 1 else r"[a-z]{0,2}"))
    patron = r"(?<![a-z0-9])" + r"\s+".join(cuerpo)
    if exacta:
        patron += r"(?![a-z0-9])"
    return re.compile(patron)


# ------------------------------------------------------------------ creación del Excel por defecto
def _cabecera(ws, titulos, anchos):
    ws.append(titulos)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1A56B0")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = "A2"


def _fuente_filas(ws):
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="top", wrap_text=True)


def crear_por_defecto(ruta: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Instrucciones"
    for linea in INSTRUCCIONES:
        ws.append([linea])
    ws.column_dimensions["A"].width = 130
    for i, fila in enumerate(ws.iter_rows(), 1):
        for c in fila:
            c.font = Font(name="Arial", size=11, bold=(i == 1))
            c.alignment = Alignment(wrap_text=True, vertical="top")

    g = wb.create_sheet("Grupos")
    _cabecera(g, ["Grupo", "Puntos", "Descripción"], [42, 10, 70])
    for fila in GRUPOS:
        g.append(list(fila))
    _fuente_filas(g)

    p = wb.create_sheet("Palabras")
    _cabecera(p, ["Grupo", "Palabra o expresión", "Tipo", "Activa"], [42, 40, 12, 10])
    for grupo, texto, tipo in PALABRAS:
        p.append([grupo, texto, tipo, "SÍ"])
    _fuente_filas(p)
    dv_g = DataValidation(type="list", formula1="=Grupos!$A$2:$A$60", allow_blank=True)
    dv_t = DataValidation(type="list", formula1='"Normal,Exacta"', allow_blank=True)
    dv_a = DataValidation(type="list", formula1='"SÍ,NO"', allow_blank=True)
    for dv, col in ((dv_g, "A"), (dv_t, "C"), (dv_a, "D")):
        p.add_data_validation(dv)
        dv.add(f"{col}2:{col}1000")

    c = wb.create_sheet("CPV")
    _cabecera(c, ["Prefijo del código CPV", "Puntos", "Descripción"], [24, 10, 50])
    for fila in CPV:
        c.append([fila[0], fila[1], fila[2]])
    for fila in c.iter_rows(min_row=2, max_col=1):
        for celda in fila:
            celda.number_format = "@"
    _fuente_filas(c)

    x = wb.create_sheet("Exclusiones")
    _cabecera(x, ["Palabra o expresión", "Tipo", "Activa"], [40, 12, 10])
    for texto, tipo in EXCLUSIONES:
        x.append([texto, tipo, "SÍ"])
    _fuente_filas(x)
    dv_t2 = DataValidation(type="list", formula1='"Normal,Exacta"', allow_blank=True)
    dv_a2 = DataValidation(type="list", formula1='"SÍ,NO"', allow_blank=True)
    x.add_data_validation(dv_t2)
    x.add_data_validation(dv_a2)
    dv_t2.add("B2:B1000")
    dv_a2.add("C2:C1000")

    a = wb.create_sheet("Ajustes")
    _cabecera(a, ["Ajuste", "Valor", "Explicación"], [58, 14, 80])
    for fila in AJUSTES + AJUSTES_EMPRESA:
        a.append(list(fila))
    _fuente_filas(a)
    for fila in a.iter_rows(min_row=2, min_col=2, max_col=2):
        for celda in fila:
            celda.font = Font(name="Arial", size=10, color="0000FF")   # azul = valor que puedes cambiar
            celda.fill = PatternFill("solid", fgColor="FFFFE0")
    wb.save(ruta)


# ------------------------------------------------------------------ lectura
def _txt(v) -> str:
    return "" if v is None else str(v).strip()


def _activa(v) -> bool:
    return norm(_txt(v)) not in ("no", "n", "0", "false", "falso", "off")


def _entero(v, defecto):
    try:
        return int(float(str(v).replace(",", ".")))
    except (TypeError, ValueError):
        return defecto


def _numero(v):
    try:
        return float(str(v).replace(".", "").replace(",", ".")) if isinstance(v, str) else (
            float(v) if v not in (None, "") else None)
    except ValueError:
        return None


def _filas(wb, hoja):
    if hoja not in wb.sheetnames:
        return None
    return [f for f in wb[hoja].iter_rows(min_row=2, values_only=True) if any(_txt(x) for x in f)]


def cargar(base: Path, log=print) -> Reglas:
    """Lee palabras_clave.xlsx (lo crea si no existe) y deja las reglas en palabras.ACTUAL."""
    global ACTUAL
    ruta = Path(base) / C.ARCHIVO_PALABRAS
    if not ruta.exists():
        crear_por_defecto(ruta)
        log(f"  Se ha creado {ruta.name} con las palabras clave por defecto: edítalo para ajustarlo a tu empresa.")
    wb = load_workbook(ruta, data_only=True)
    r = Reglas(ruta=ruta)

    for f in _filas(wb, "Grupos") or []:
        if _txt(f[0]):
            r.grupos[_txt(f[0])] = _entero(f[1] if len(f) > 1 else None, 0)

    for i, f in enumerate(_filas(wb, "Palabras") or [], 2):
        grupo, texto = _txt(f[0]), _txt(f[1] if len(f) > 1 else "")
        if not texto or not _activa(f[3] if len(f) > 3 else None):
            continue
        if grupo not in r.grupos:
            r.avisos.append(f"Hoja Palabras, fila {i}: el grupo «{grupo}» no existe en la hoja Grupos; se ignora «{texto}».")
            continue
        exacta = norm(_txt(f[2] if len(f) > 2 else "")) == "exacta"
        rx = compilar(texto, exacta)
        if rx:
            r.terminos.append(Termino(grupo, texto, rx, exacta))

    for f in _filas(wb, "CPV") or []:
        pref = _txt(f[0]).replace(" ", "")
        if pref.endswith(".0"):
            pref = pref[:-2]
        if pref:
            r.cpv[pref] = (_entero(f[1] if len(f) > 1 else None, 0), _txt(f[2] if len(f) > 2 else ""))

    for f in _filas(wb, "Exclusiones") or []:
        texto = _txt(f[0])
        if texto and _activa(f[2] if len(f) > 2 else None):
            exacta = norm(_txt(f[1] if len(f) > 1 else "")) == "exacta"
            rx = compilar(texto, exacta)
            if rx:
                r.exclusiones.append(Termino("Exclusión", texto, rx, exacta))

    ajustes = {norm(_txt(f[0])): f[1] if len(f) > 1 else None for f in (_filas(wb, "Ajustes") or [])}

    def aj(clave, defecto, conv=_entero):
        for k, v in ajustes.items():
            if k.startswith(norm(clave)):
                return conv(v, defecto) if conv is _entero else conv(v)
        return defecto

    r.minima = aj("Puntuación mínima para que aparezca", C.PUNTUACION_MINIMA)
    r.alto = aj("Puntuación a partir de la cual", C.PUNTUACION_ENCAJA)
    r.tope = aj("Tope de puntos", C.TOPE_PALABRAS)
    r.minima_lote = aj("Puntuación mínima para considerar un lote", 20)
    r.penalizacion = aj("Penalización por palabra", 40)
    r.importe_min = aj("Importe mínimo", None, _numero)
    r.importe_max = aj("Importe máximo", None, _numero)
    r.empresa = aj("Nombre de la empresa", "", _txt)
    r.logo_claro = aj("Logo para el tema claro", "", _txt)
    r.logo_oscuro = aj("Logo para el tema oscuro", "", _txt)
    carpeta_logos = Path(base) / "web" / "config"
    for campo in ("logo_claro", "logo_oscuro"):
        nombre = getattr(r, campo)
        if nombre and (Path(nombre).name != nombre or not (carpeta_logos / nombre).is_file()):
            r.avisos.append(f"Ajustes: el logo «{nombre}» no está en web/config; se usa el icono por defecto.")
            setattr(r, campo, "")
    if not r.terminos:
        r.avisos.append("No hay ninguna palabra clave activa: no se encontrará nada. Revisa la hoja Palabras.")
    ACTUAL = r
    return r
