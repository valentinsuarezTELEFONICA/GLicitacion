# -*- coding: utf-8 -*-
"""Comunidad autónoma, provincia, ciudad y país a partir de códigos NUTS, nombres o texto libre."""
from __future__ import annotations

import re

from textos import norm

NUTS2 = {
    "ES11": "Galicia", "ES12": "Asturias", "ES13": "Cantabria", "ES21": "País Vasco",
    "ES22": "Navarra", "ES23": "La Rioja", "ES24": "Aragón", "ES30": "Madrid",
    "ES41": "Castilla y León", "ES42": "Castilla-La Mancha", "ES43": "Extremadura",
    "ES51": "Cataluña", "ES52": "Comunidad Valenciana", "ES53": "Baleares", "ES61": "Andalucía",
    "ES62": "Murcia", "ES63": "Ceuta", "ES64": "Melilla", "ES70": "Canarias",
}

NUTS3 = {
    "ES111": "A Coruña", "ES112": "Lugo", "ES113": "Ourense", "ES114": "Pontevedra",
    "ES120": "Asturias", "ES130": "Cantabria", "ES211": "Álava", "ES212": "Gipuzkoa",
    "ES213": "Bizkaia", "ES220": "Navarra", "ES230": "La Rioja", "ES241": "Huesca",
    "ES242": "Teruel", "ES243": "Zaragoza", "ES300": "Madrid", "ES411": "Ávila",
    "ES412": "Burgos", "ES413": "León", "ES414": "Palencia", "ES415": "Salamanca",
    "ES416": "Segovia", "ES417": "Soria", "ES418": "Valladolid", "ES419": "Zamora",
    "ES421": "Albacete", "ES422": "Ciudad Real", "ES423": "Cuenca", "ES424": "Guadalajara",
    "ES425": "Toledo", "ES431": "Badajoz", "ES432": "Cáceres", "ES511": "Barcelona",
    "ES512": "Girona", "ES513": "Lleida", "ES514": "Tarragona", "ES521": "Alicante",
    "ES522": "Castellón", "ES523": "Valencia", "ES531": "Ibiza y Formentera",
    "ES532": "Mallorca", "ES533": "Menorca", "ES611": "Almería", "ES612": "Cádiz",
    "ES613": "Córdoba", "ES614": "Granada", "ES615": "Huelva", "ES616": "Jaén",
    "ES617": "Málaga", "ES618": "Sevilla", "ES620": "Murcia", "ES630": "Ceuta",
    "ES640": "Melilla", "ES703": "El Hierro", "ES704": "Fuerteventura", "ES705": "Gran Canaria",
    "ES706": "La Gomera", "ES707": "La Palma", "ES708": "Lanzarote", "ES709": "Tenerife",
}

# Nombre (sin tildes, minúsculas) -> (provincia o isla, comunidad autónoma). "" = solo comunidad.
LUGARES = {
    "a coruna": ("A Coruña", "Galicia"), "la coruna": ("A Coruña", "Galicia"), "coruna": ("A Coruña", "Galicia"),
    "lugo": ("Lugo", "Galicia"), "ourense": ("Ourense", "Galicia"), "orense": ("Ourense", "Galicia"),
    "pontevedra": ("Pontevedra", "Galicia"), "galicia": ("", "Galicia"),
    "asturias": ("Asturias", "Asturias"), "cantabria": ("Cantabria", "Cantabria"),
    "alava": ("Álava", "País Vasco"), "araba": ("Álava", "País Vasco"),
    "gipuzkoa": ("Gipuzkoa", "País Vasco"), "guipuzcoa": ("Gipuzkoa", "País Vasco"),
    "bizkaia": ("Bizkaia", "País Vasco"), "vizcaya": ("Bizkaia", "País Vasco"),
    "pais vasco": ("", "País Vasco"), "euskadi": ("", "País Vasco"),
    "navarra": ("Navarra", "Navarra"), "la rioja": ("La Rioja", "La Rioja"), "rioja": ("La Rioja", "La Rioja"),
    "huesca": ("Huesca", "Aragón"), "teruel": ("Teruel", "Aragón"), "zaragoza": ("Zaragoza", "Aragón"),
    "aragon": ("", "Aragón"), "madrid": ("Madrid", "Madrid"),
    "avila": ("Ávila", "Castilla y León"), "burgos": ("Burgos", "Castilla y León"),
    "leon": ("León", "Castilla y León"), "palencia": ("Palencia", "Castilla y León"),
    "salamanca": ("Salamanca", "Castilla y León"), "segovia": ("Segovia", "Castilla y León"),
    "soria": ("Soria", "Castilla y León"), "valladolid": ("Valladolid", "Castilla y León"),
    "zamora": ("Zamora", "Castilla y León"), "castilla y leon": ("", "Castilla y León"),
    "albacete": ("Albacete", "Castilla-La Mancha"), "ciudad real": ("Ciudad Real", "Castilla-La Mancha"),
    "cuenca": ("Cuenca", "Castilla-La Mancha"), "guadalajara": ("Guadalajara", "Castilla-La Mancha"),
    "toledo": ("Toledo", "Castilla-La Mancha"), "castilla-la mancha": ("", "Castilla-La Mancha"),
    "castilla la mancha": ("", "Castilla-La Mancha"),
    "badajoz": ("Badajoz", "Extremadura"), "caceres": ("Cáceres", "Extremadura"),
    "extremadura": ("", "Extremadura"),
    "barcelona": ("Barcelona", "Cataluña"), "girona": ("Girona", "Cataluña"), "gerona": ("Girona", "Cataluña"),
    "lleida": ("Lleida", "Cataluña"), "lerida": ("Lleida", "Cataluña"), "tarragona": ("Tarragona", "Cataluña"),
    "cataluna": ("", "Cataluña"), "catalunya": ("", "Cataluña"),
    "alicante": ("Alicante", "Comunidad Valenciana"), "alacant": ("Alicante", "Comunidad Valenciana"),
    "castellon": ("Castellón", "Comunidad Valenciana"), "castello": ("Castellón", "Comunidad Valenciana"),
    "valencia": ("Valencia", "Comunidad Valenciana"),
    "comunidad valenciana": ("", "Comunidad Valenciana"), "comunitat valenciana": ("", "Comunidad Valenciana"),
    "generalitat valenciana": ("", "Comunidad Valenciana"),
    "mallorca": ("Mallorca", "Baleares"), "menorca": ("Menorca", "Baleares"), "ibiza": ("Ibiza y Formentera", "Baleares"),
    "eivissa": ("Ibiza y Formentera", "Baleares"), "formentera": ("Ibiza y Formentera", "Baleares"),
    "illes balears": ("", "Baleares"), "islas baleares": ("", "Baleares"), "baleares": ("", "Baleares"),
    "almeria": ("Almería", "Andalucía"), "cadiz": ("Cádiz", "Andalucía"), "cordoba": ("Córdoba", "Andalucía"),
    "granada": ("Granada", "Andalucía"), "huelva": ("Huelva", "Andalucía"), "jaen": ("Jaén", "Andalucía"),
    "malaga": ("Málaga", "Andalucía"), "sevilla": ("Sevilla", "Andalucía"), "andalucia": ("", "Andalucía"),
    "murcia": ("Murcia", "Murcia"), "ceuta": ("Ceuta", "Ceuta"), "melilla": ("Melilla", "Melilla"),
    "las palmas": ("Gran Canaria", "Canarias"), "santa cruz de tenerife": ("Tenerife", "Canarias"),
    "tenerife": ("Tenerife", "Canarias"), "gran canaria": ("Gran Canaria", "Canarias"),
    "lanzarote": ("Lanzarote", "Canarias"), "fuerteventura": ("Fuerteventura", "Canarias"),
    "la palma": ("La Palma", "Canarias"), "la gomera": ("La Gomera", "Canarias"),
    "el hierro": ("El Hierro", "Canarias"), "canarias": ("", "Canarias"),
}

PAISES = {"ES": "España", "ESP": "España", "PT": "Portugal", "PRT": "Portugal", "FR": "Francia", "FRA": "Francia",
          "DE": "Alemania", "DEU": "Alemania", "IT": "Italia", "ITA": "Italia", "AD": "Andorra", "AND": "Andorra"}

_ORDEN = sorted(LUGARES, key=len, reverse=True)
_RE_LUGARES = re.compile(r"(?<![a-z0-9])(" + "|".join(re.escape(k) for k in _ORDEN) + r")(?![a-z0-9])")


def nombre_pais(codigo_o_nombre: str) -> str:
    t = (codigo_o_nombre or "").strip()
    return PAISES.get(t.upper(), t)


def buscar_en_texto(texto: str):
    """Primer lugar conocido que aparece en un texto: (provincia, comunidad) o None."""
    m = _RE_LUGARES.search(norm(texto))
    return LUGARES[m.group(1)] if m else None


def resolver(nuts: str = "", provincia: str = "", ciudad: str = "", pais: str = "", texto: str = "") -> dict:
    """Devuelve pais, ccaa, provincia, ciudad y de dónde se ha sacado ('dato', 'NUTS', 'inferido')."""
    res = {"pais": nombre_pais(pais), "ccaa": "", "provincia": "", "ciudad": (ciudad or "").strip(), "origen": ""}
    nuts = (nuts or "").strip().upper()
    if nuts.startswith("ES"):
        res["pais"] = res["pais"] or "España"
        if nuts[:4] in NUTS2:
            res["ccaa"] = NUTS2[nuts[:4]]
            res["origen"] = "NUTS"
        if nuts[:5] in NUTS3:
            res["provincia"] = NUTS3[nuts[:5]]
    prov = (provincia or "").strip()
    if prov:
        hit = LUGARES.get(norm(prov))
        if hit:
            res["provincia"] = res["provincia"] or hit[0]
            res["ccaa"] = res["ccaa"] or hit[1]
            res["origen"] = res["origen"] or "dato"
        elif not res["provincia"]:
            res["provincia"] = prov
    if not res["ccaa"] and res["ciudad"]:
        hit = LUGARES.get(norm(res["ciudad"]))
        if hit:
            res["provincia"] = res["provincia"] or hit[0]
            res["ccaa"] = hit[1]
            res["origen"] = "dato"
    if not res["ccaa"] and texto:
        hit = buscar_en_texto(texto)
        if hit:
            res["provincia"] = res["provincia"] or hit[0]
            res["ccaa"] = hit[1]
            res["origen"] = "inferido"
    if res["ccaa"] and not res["pais"]:
        res["pais"] = "España"
    return res
