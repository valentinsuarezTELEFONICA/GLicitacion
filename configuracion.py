# -*- coding: utf-8 -*-
"""
CONFIGURACIÓN TÉCNICA.

Las palabras clave, los CPV, las exclusiones y los umbrales de puntuación ya NO están aquí:
se editan en el Excel «palabras_clave.xlsx» (se crea solo la primera vez).
"""

EMPRESA = "Mi empresa"

# Descripción de la empresa. Solo se usa si activas el resumen con IA (opcional).
PERFIL_EMPRESA = (
    "Empresa tecnológica que presenta ofertas a licitaciones públicas de: realidad virtual, "
    "realidad aumentada, experiencias inmersivas, museos y centros de interpretación, hardware "
    "informático, equipamiento audiovisual, videowalls y pantallas interactivas, digitalización, "
    "software a medida, gestores documentales, gestores de contenidos (CMS), aplicaciones web y "
    "móviles, gamificación y tecnología para educación, cultura y turismo."
)

# ---------------------------------------------------------------- Ventana de tiempo
DIAS_ATRAS_POR_DEFECTO = 1    # primera ejecución, o cuando usas --dias
MAX_DIAS_ATRAS = 7            # tope si pasas varios días sin ejecutarlo

# ---------------------------------------------------------------- Fuentes
FUENTES_ACTIVAS = ["placsp", "boe", "ted", "extra"]

# Plataforma de Contratación del Sector Público (feed oficial de datos abiertos)
PLACSP_FEED = ("https://contrataciondelsectorpublico.gob.es/sindicacion/sindicacion_643/"
               "licitacionesPerfilesContratanteCompleto3.atom")
PLACSP_MAX_PAGINAS = 40          # cada página trae hasta 500 anuncios
ESTADOS_VALIDOS = {"PUB"}        # PUB = publicada (plazo abierto)

BOE_API = "https://www.boe.es/datosabiertos/api/boe/sumario/{fecha}"
BOE_LEER_TEXTO = True            # abre el anuncio del BOE para sacar plazo, importe y lugar
TED_API = "https://api.ted.europa.eu/v3/notices/search"
TED_MAX_PAGINAS = 10

# Campos que se piden a TED. Los «extra» se prueban uno a uno: los que TED no reconozca se
# descartan solos y el resultado se recuerda para las siguientes ejecuciones.
TED_CAMPOS_BASE = ["publication-number", "notice-title", "buyer-name", "classification-cpv",
                   "deadline-receipt-tender-date-lot", "publication-date", "links"]
TED_CAMPOS_EXTRA = ["notice-type", "buyer-city", "buyer-country", "place-of-performance-city-lot",
                    "place-of-performance-country-lot", "estimated-value-lot", "estimated-value-cur-lot",
                    "procedure-type", "contract-nature", "title-lot", "description-lot"]

# Otros sitios con RSS o Atom (boletines autonómicos, diputaciones, universidades...).
# La mayoría de organismos ya salen en la Plataforma de Contratación; añade aquí solo los
# que tengan plataforma propia. Formato: ("Nombre que verás en el informe", "URL del feed")
FEEDS_EXTRA = [
    # ("Mi comunidad autónoma", "https://.../licitaciones.rss"),
]

# ---------------------------------------------------------------- Palabras clave (Excel)
ARCHIVO_PALABRAS = "palabras_clave.xlsx"
# Valores de respaldo si falta algún ajuste en el Excel
PUNTUACION_MINIMA = 25
PUNTUACION_ENCAJA = 50
TOPE_PALABRAS = 70

# ---------------------------------------------------------------- Análisis de pliegos
DESCARGAR_PLIEGOS = True
MAX_LICITACIONES_ANALIZADAS = 30     # solo se descargan pliegos de las mejor puntuadas
MAX_DOCS_POR_LICITACION = 3          # PCAP, PPT y un anexo
MAX_MB_POR_DOCUMENTO = 60
LINKS_LOCALES = True                 # los enlaces a página abren la copia descargada

# Qué buscar dentro de los pliegos. Concepto -> patrones (sin tildes, minúsculas).
CONCEPTOS = {
    "Presupuesto / valor estimado": [r"valor estimado", r"presupuesto base de licitacion"],
    "Presentación de ofertas": [r"plazo de presentacion", r"presentacion de (ofertas|proposiciones)",
                                r"\bdeuc\b", r"sobre (a|b|c|1|2|3|unico)\b"],
    "Solvencia económica": [r"solvencia economica", r"volumen anual de negocios",
                            r"cifra anual de negocios"],
    "Solvencia técnica": [r"solvencia tecnica", r"trabajos (realizados|similares)",
                          r"(servicios|suministros) similares", r"relacion de (los )?principales"],
    "Clasificación empresarial": [r"clasificacion (empresarial|del contratista)",
                                  r"grupo [a-z],? subgrupo", r"categoria \d"],
    "Garantías": [r"garantia (definitiva|provisional|complementaria)", r"garantia del \d"],
    "Certificaciones": [r"\biso ?\d{4,5}", r"\biso/iec", r"esquema nacional de (seguridad|interoperabilidad)",
                        r"certificad[oa] (de|en) "],
    "Criterios de adjudicación": [r"criterios de (adjudicacion|valoracion)", r"juicio de valor",
                                  r"criterios (automaticos|evaluables mediante formulas)"],
    "Plazo de ejecución": [r"plazo de (ejecucion|entrega|duracion)", r"duracion del contrato"],
    "Equipo y medios personales": [r"equipo (de trabajo|tecnico|minimo)", r"medios personales",
                                   r"adscripcion de medios", r"titulacion"],
    "Penalidades y riesgos": [r"penalidad", r"penalizacion", r"resolucion del contrato",
                              r"propiedad intelectual", r"subcontrata"],
}
PAGINAS_POR_CONCEPTO = 5   # páginas que se enlazan por concepto y documento

# ---------------------------------------------------------------- Resumen con IA (OPCIONAL)
# Solo funciona si defines la variable de entorno ANTHROPIC_API_KEY. Sin ella, el informe
# usa extractos automáticos del pliego. Con ella, además redacta un resumen de lo que se pide.
USAR_IA = True
MODELO_IA = "claude-sonnet-5-5"

# ---------------------------------------------------------------- Salida
#
# Carpeta donde se dejan los informes Y la web de los comerciales (index.html, datos y pliegos).
# Para compartirla, apúntala a una carpeta de OneDrive con permisos para los comerciales, por ejemplo:
#     CARPETA_INFORMES = r"C:\Users\valen\OneDrive - Mi empresa\Licitaciones"
# Si es relativa, se crea junto al programa. Dentro hay una carpeta por día de búsqueda.
CARPETA_INFORMES = "informes"

# Genera/actualiza la web (index.html + carpeta datos) en cada ejecución. Los comerciales solo la abren.
WEB_ACTIVA = True
