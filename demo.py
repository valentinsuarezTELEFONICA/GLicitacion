# -*- coding: utf-8 -*-
"""Datos de demostración (sin conexión) para ver cómo queda el informe."""
from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate

from analisis import Documento, Licitacion, Lote

PAGINAS_PCAP = [
    ("1. Objeto del contrato", "El presente pliego tiene por objeto la contratación del diseño, producción e instalación "
     "de experiencias inmersivas de realidad virtual para el Centro de Interpretación. El valor estimado del contrato asciende a 185.000 euros."),
    ("2. Régimen jurídico", "El contrato se rige por la Ley 9/2017 de Contratos del Sector Público. Los licitadores podrán "
     "presentar sus ofertas en el plazo de presentación indicado en el anuncio, mediante los sobres A y B."),
    ("3. Solvencia económica y financiera", "La solvencia económica se acreditará mediante el volumen anual de negocios en el "
     "mejor ejercicio de los tres últimos, que deberá ser al menos 1,5 veces el valor estimado del contrato."),
    ("4. Solvencia técnica", "La solvencia técnica se acreditará mediante una relación de los trabajos realizados en los últimos "
     "cinco años de naturaleza similar, con importe acumulado mínimo de 90.000 euros."),
    ("5. Garantías", "El licitador propuesto como adjudicatario deberá constituir una garantía definitiva del 5% del precio de "
     "adjudicación, IVA excluido."),
    ("6. Certificaciones", "Se exigirá estar en posesión de la certificación ISO 9001 y del certificado de conformidad con el "
     "Esquema Nacional de Seguridad en categoría media."),
    ("7. Criterios de adjudicación", "Criterios evaluables mediante juicio de valor: hasta 40 puntos. Criterios evaluables mediante "
     "fórmulas: oferta económica hasta 45 puntos y mejoras hasta 15 puntos."),
    ("8. Plazo y penalidades", "El plazo de ejecución será de seis meses desde la formalización. Por retrasos se aplicarán "
     "penalidades de 0,60 euros por cada 1.000 euros del precio del contrato. La propiedad intelectual corresponde al órgano de contratación."),
]


def _pdf(ruta: Path, titulo: str, paginas) -> None:
    est = getSampleStyleSheet()
    elementos = []
    for i, (h, t) in enumerate(paginas):
        elementos += [Paragraph(titulo, est["Title"]), Paragraph(h, est["Heading2"]), Paragraph(t, est["BodyText"])]
        if i < len(paginas) - 1:
            elementos.append(PageBreak())
    SimpleDocTemplate(str(ruta), pagesize=A4).build(elementos)


def datos_demo(carpeta: Path) -> list[Licitacion]:
    carpeta.mkdir(parents=True, exist_ok=True)
    pcap, ppt, pantallas, boe = (carpeta / n for n in ("pcap.pdf", "ppt.pdf", "pantallas.pdf", "boe.pdf"))
    _pdf(pcap, "Pliego de cláusulas administrativas particulares (demo)", PAGINAS_PCAP)
    _pdf(ppt, "Pliego de prescripciones técnicas (demo)", [
        ("1. Alcance", "Se requiere equipo de trabajo mínimo de un director de proyecto, dos desarrolladores 3D y un técnico audiovisual."),
        ("2. Requisitos técnicos", "Los visores serán autónomos con seguimiento de manos. Plazo de entrega: tres meses.")])
    _pdf(pantallas, "Pliego (demo)", PAGINAS_PCAP[:6])
    _pdf(boe, "BOE-B-2026-99999 (demo)", [("Anuncio de licitación", "Objeto: desarrollo de aplicación móvil y portal web de turismo. "
         "Presupuesto base de licitación: 62.000 euros. Plazo de presentación de ofertas: 15 días naturales.")])
    u = lambda p: p.resolve().as_uri()
    ahora = datetime.now()
    base = [
        Licitacion(id="demo-1", fuente="PLACSP", expediente="DEMO-2026/001",
                   titulo="Servicio de diseño, producción e instalación de experiencias inmersivas de realidad virtual para el Centro de Interpretación del Patrimonio",
                   organo="Ayuntamiento de Ejemplo (demo)", tipo="Servicios", procedimiento="Abierto", duracion="6 meses",
                   presupuesto=185000, valor_estimado=222000, fecha_limite=ahora + timedelta(days=9),
                   cpv=["92521000", "48000000", "90910000"], estado="PUB",
                   url="https://contrataciondelsectorpublico.gob.es/", actualizada=ahora, fecha_publicacion=ahora,
                   pais="España", ccaa="Andalucía", provincia="Sevilla", ciudad="Sevilla (demo)", ubicacion_origen="dato",
                   lotes=[Lote("1", "Diseño y desarrollo de experiencias inmersivas de realidad virtual", "", 120000, ["72212000"]),
                          Lote("2", "Suministro de visores y equipamiento informático", "", 45000, ["30213000"]),
                          Lote("3", "Servicio de limpieza y mantenimiento de las instalaciones", "", 20000, ["90910000"])],
                   documentos=[Documento("PCAP", "pcap.pdf", u(pcap)), Documento("PPT", "ppt.pdf", u(ppt))]),
        Licitacion(id="demo-2", fuente="PLACSP", expediente="DEMO-2026/014",
                   titulo="Suministro e instalación de videowall y pantallas interactivas para el Museo Provincial",
                   organo="Diputación de Ejemplo (demo)", tipo="Suministros", procedimiento="Abierto simplificado",
                   presupuesto=96500, valor_estimado=96500, fecha_limite=ahora + timedelta(days=4), cpv=["32321200"],
                   estado="PUB", url="https://contrataciondelsectorpublico.gob.es/", actualizada=ahora, fecha_publicacion=ahora,
                   pais="España", ccaa="Madrid", provincia="Madrid", ciudad="Alcalá de Henares (demo)", ubicacion_origen="dato",
                   documentos=[Documento("PCAP", "pantallas.pdf", u(pantallas))]),
        Licitacion(id="BOE-B-2026-99999", fuente="BOE", expediente="TUR-2026/7",
                   titulo="Anuncio de licitación de: Consorcio de Turismo de Ejemplo (Valencia). Objeto: Desarrollo de aplicación móvil y portal web de turismo. Expediente: TUR-2026/7.",
                   organo="Consorcio de Turismo de Ejemplo (Valencia)", url="https://www.boe.es/", estado="PUB", actualizada=ahora,
                   fecha_publicacion=ahora, ccaa="Comunidad Valenciana", provincia="Valencia", pais="España", ubicacion_origen="inferido",
                   documentos=[Documento("Anuncio BOE", "BOE-B-2026-99999.pdf", u(boe))]),
        Licitacion(id="demo-4", fuente="PLACSP", expediente="DEMO-2026/020",
                   titulo="Servicio de limpieza de edificios municipales", organo="Ayuntamiento de Ejemplo (demo)",
                   tipo="Servicios", presupuesto=300000, fecha_limite=ahora + timedelta(days=12), cpv=["90910000"],
                   estado="PUB", url="https://contrataciondelsectorpublico.gob.es/", actualizada=ahora),
    ]
    return base + _extras(ahora, pcap, ppt, u)


# (título, órgano, comunidad, provincia, ciudad, presupuesto, días hasta el plazo, CPV, días atrás en que se recogió, fuente)
EXTRAS = [
    ("Desarrollo de aplicación de realidad aumentada para la ruta turística del casco histórico",
     "Ayuntamiento de Toledo", "Castilla-La Mancha", "Toledo", "Toledo", 78000, 14, "72212000", 1, "PLACSP"),
    ("Suministro de pantallas táctiles interactivas y tótems para el centro de visitantes",
     "Diputación de Girona", "Cataluña", "Girona", "Girona", 124000, 6, "32321200", 1, "PLACSP"),
    ("Digitalización de archivo histórico y suministro de gestor documental",
     "Gobierno de Navarra", "Navarra", "Navarra", "Pamplona", 210000, 20, "48311000", 2, "PLACSP"),
    ("Plataforma e-learning y aula virtual para la formación del profesorado",
     "Consejería de Educación de Galicia", "Galicia", "A Coruña", "Santiago de Compostela", 95000, 11, "72212000", 2, "PLACSP"),
    ("Equipamiento audiovisual y proyección para el auditorio municipal",
     "Ayuntamiento de Zaragoza", "Aragón", "Zaragoza", "Zaragoza", 68000, 3, "32320000", 0, "PLACSP"),
    ("Museografía y audiovisuales inmersivos para el museo de la ciencia",
     "Junta de Andalucía", "Andalucía", "Málaga", "Málaga", 340000, 25, "92521000", 0, "PLACSP"),
    ("Mantenimiento de ordenadores y servidores del servicio de informática",
     "Universidad de Murcia", "Murcia", "Murcia", "Murcia", 52000, 9, "72250000", 3, "PLACSP"),
    ("Desarrollo de portal web y gestor de contenidos CMS para la oficina de turismo",
     "Cabildo de Tenerife", "Canarias", "Tenerife", "Santa Cruz de Tenerife", 56000, 17, "72413000", 0, "PLACSP"),
    ("Experiencia inmersiva y videomapping para el centro de interpretación del patrimonio",
     "Principado de Asturias", "Asturias", "Asturias", "Oviedo", 145000, 4, "92521000", 1, "PLACSP"),
    ("Videojuego educativo y gamificación para el programa de patrimonio escolar",
     "Generalitat Valenciana", "Comunidad Valenciana", "Valencia", "Valencia", 61000, 22, "48000000", 4, "PLACSP"),
    ("Suministro de tablets y equipamiento informático para centros educativos",
     "Comunidad de Madrid", "Madrid", "Madrid", "Madrid", 420000, 13, "30213000", 0, "PLACSP"),
    ("Creación de audioguías y señalización digital para el parque arqueológico",
     "Junta de Castilla y León", "Castilla y León", "Burgos", "Burgos", 83000, 7, "92520000", 2, "PLACSP"),
    ("Desarrollo de app móvil de turismo y portal web de la isla",
     "Consell de Mallorca", "Baleares", "Mallorca", "Palma", 64000, 18, "72212000", 1, "PLACSP"),
    ("Realidad virtual para el museo marítimo y su exposición permanente",
     "Diputación Foral de Bizkaia", "País Vasco", "Bizkaia", "Bilbao", 99000, 8, "92521000", 3, "PLACSP"),
    ("Sistema de realidad aumentada para la visita al yacimiento romano",
     "Junta de Extremadura", "Extremadura", "Badajoz", "Mérida", 71000, -2, "72212000", 5, "PLACSP"),
    ("Servicio de desarrollo de aplicación móvil y portal web para la universidad",
     "Universidad de Zaragoza", "Aragón", "Zaragoza", "Zaragoza", 90000, 28, "72212000", 0, "TED"),
]


def _extras(ahora, pcap, ppt, u) -> list:
    out = []
    for n, (titulo, organo, ccaa, prov, ciudad, presup, plazo, cpv, atras, fuente) in enumerate(EXTRAS, 1):
        lic = Licitacion(
            id=f"demo-e{n}", fuente=fuente, expediente=f"DEMO-2026/1{n:02d}", titulo=titulo,
            organo=organo + " (demo)", tipo="Suministros" if "Suministro" in titulo else "Servicios",
            procedimiento=["Abierto", "Abierto simplificado", "Restringido"][n % 3],
            duracion=f"{6 + (n % 4) * 6} meses", presupuesto=presup, valor_estimado=round(presup * 1.2),
            fecha_limite=ahora + timedelta(days=plazo), cpv=[cpv], estado="PUB",
            url="https://contrataciondelsectorpublico.gob.es/", actualizada=ahora, fecha_publicacion=ahora,
            pais="España", ccaa=ccaa, provincia=prov, ciudad=ciudad, ubicacion_origen="dato",
            documentos=[Documento("PCAP", "pcap.pdf", u(pcap)), Documento("PPT", "ppt.pdf", u(ppt))] if n % 3 else [])
        lic.dias_atras_demo = atras
        if "museo de la ciencia" in titulo:
            lic.lotes = [Lote("1", "Museografía e instalaciones audiovisuales inmersivas", "", 210000, ["92521000"]),
                         Lote("2", "Suministro de equipamiento informático y pantallas interactivas", "", 90000, ["30231000"]),
                         Lote("3", "Obra de reforma de la sala y mobiliario", "", 40000, ["45210000"])]
        out.append(lic)
    return out
