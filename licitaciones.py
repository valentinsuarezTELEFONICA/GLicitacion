#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Licitaciones diarias: lee las fuentes oficiales, puntúa lo que podría encajar con tu empresa
(según palabras_clave.xlsx), analiza los pliegos (con página) y genera un Excel y un PDF por día.
Recuerda lo ya recogido para no duplicarlo y avisa de los cambios.

Uso:
    py licitaciones.py              # revisión de hoy
    py licitaciones.py --demo       # ejemplo sin conexión, para ver el aspecto del informe
    py licitaciones.py --dias 3     # revisar los últimos 3 días
    py licitaciones.py --abrir      # abre el Excel al terminar
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import configuracion as C
import exportar_web
import fichas as F
import fuentes
import informes
import palabras
from analisis import analizar_pdf, ia_disponible, puntuar, resumen_ia
from registro import Registro

BASE = Path(__file__).resolve().parent
PRIORIDAD = {"PLACSP": 0, "TED": 1, "BOE": 2}      # si una licitación sale en varias fuentes, manda la primera


def periodo(args, registro: Registro) -> datetime:
    hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    ultima = registro.meta.get("ultima_ejecucion")
    if args.dias:
        dias = args.dias
    elif ultima:
        dias = max(1, (hoy - datetime.fromisoformat(ultima).replace(hour=0, minute=0, second=0)).days)
    else:
        dias = C.DIAS_ATRAS_POR_DEFECTO
    return hoy - timedelta(days=min(dias, C.MAX_DIAS_ATRAS))


def recoger(desde: datetime, activas: list, registro: Registro):
    """Devuelve (licitaciones, estado_de_fuentes). Una fuente caída no detiene el resto."""
    trabajos = []
    if "placsp" in activas:
        trabajos.append(("PLACSP (Plataforma de Contratación)", lambda log: fuentes.leer_placsp(desde, log)))
    if "boe" in activas:
        trabajos.append(("BOE (sección V-A)", lambda log: fuentes.leer_boe(desde, log)))
    if "ted" in activas:
        trabajos.append(("TED (Europa, España)", lambda log: fuentes.leer_ted(desde, log, registro.meta)))
    if "extra" in activas:
        for nombre, url in C.FEEDS_EXTRA:
            trabajos.append((nombre, lambda log, n=nombre, u=url: fuentes.leer_feed_extra(n, u, desde, log)))
    todas, estados = [], []
    for nombre, fn in trabajos:
        print(f"  · {nombre}")
        try:
            lote = fn(lambda m: print(m))
            todas += lote
            estados.append((nombre, True, f"{len(lote)} anuncios leídos"))
        except Exception as e:  # noqa: BLE001
            print(f"    ERROR: {type(e).__name__}: {e}")
            estados.append((nombre, False, f"No se pudo leer: {type(e).__name__}: {str(e)[:160]}"))
    return todas, estados


def analizar(lics: list, carpeta_pliegos: Path, sin_pliegos: bool, sin_ia: bool) -> None:
    usar_ia = ia_disponible() and not sin_ia
    for n, lic in enumerate(lics, 1):
        print(f"  [{n}/{len(lics)}] {lic.titulo[:80]}")
        if not sin_pliegos and C.DESCARGAR_PLIEGOS:
            prioridad = {"PCAP": 0, "PPT": 1}
            docs = sorted(lic.documentos, key=lambda d: prioridad.get(d.tipo, 2))[:C.MAX_DOCS_POR_LICITACION]
            sub = carpeta_pliegos / re.sub(r"[^\w\-]+", "_", lic.expediente or lic.id)[:60]
            for d in docs:
                if not d.url:
                    continue
                fuentes.descargar(d, sub)
                time.sleep(0.4)
                if d.ruta:
                    try:
                        a = analizar_pdf(d.ruta)
                        a.update(tipo=d.tipo, nombre=Path(d.ruta).name, ruta=d.ruta, url=d.url)
                        lic.analisis.append(a)
                        if not a["texto_util"]:
                            d.error = "PDF sin texto (escaneado): ábrelo a mano"
                    except Exception as e:  # noqa: BLE001
                        d.error = f"no se pudo leer el PDF ({type(e).__name__})"
        if not lic.documentos:
            lic.notas.append("El anuncio no incluye enlaces a pliegos: consúltalo en la web del anuncio.")
        if usar_ia and lic.analisis:
            try:
                lic.ia = resumen_ia(lic)
            except Exception as e:  # noqa: BLE001
                lic.notas.append(f"Resumen con IA no disponible ({type(e).__name__})")


# ------------------------------------------------------------------ pliegos que faltan
def reintentar_pliegos(registro: Registro, raiz: Path, hoy: date) -> tuple[int, int]:
    """Vuelve a descargar y analizar los pliegos que no se pudieron bajar en ejecuciones anteriores.
    Solo licitaciones con el plazo abierto (o sin plazo), las de más puntos primero.
    Devuelve (licitaciones revisadas, documentos conseguidos)."""
    from types import SimpleNamespace
    prioridad = {"PCAP": 0, "PPT": 1}
    ahora = datetime.now()
    candidatas = []
    for f in registro.fichas():
        limite = F.fecha(f.get("fecha_limite", ""))
        if limite and limite < ahora:
            continue
        docs = sorted(f.get("documentos", []), key=lambda d: prioridad.get(d.get("tipo"), 2))[:C.MAX_DOCS_POR_LICITACION]
        if any(d.get("url") and not d.get("ruta_rel") for d in docs):
            candidatas.append((f, docs))
    candidatas.sort(key=lambda x: -x[0].get("puntos", 0))
    candidatas = candidatas[:C.MAX_LICITACIONES_ANALIZADAS]
    conseguidos = 0
    for n, (f, docs) in enumerate(candidatas, 1):
        print(f"  [{n}/{len(candidatas)}] {f.get('titulo', '')[:80]}")
        sub = raiz / f.get("dia", hoy.isoformat()) / "pliegos" / re.sub(r"[^\w\-]+", "_", f.get("expediente") or f.get("id", ""))[:60]
        analisis = []
        for d in docs:
            if not d.get("url"):
                continue
            ruta_previa = raiz / d["ruta_rel"] if d.get("ruta_rel") else None
            if ruta_previa and ruta_previa.exists():
                ruta = str(ruta_previa)
            else:
                doc = fuentes.Documento(tipo=d.get("tipo", ""), nombre=d.get("nombre", ""), url=d["url"])
                fuentes.descargar(doc, sub)
                time.sleep(0.4)
                d["error"] = doc.error
                if not doc.ruta:
                    continue
                ruta = doc.ruta
                d["ruta_rel"] = F.rel(ruta, raiz)
                conseguidos += 1
            try:
                a = analizar_pdf(ruta)
                a.update(tipo=d.get("tipo", ""), nombre=Path(ruta).name, ruta=ruta, url=d["url"])
                analisis.append(a)
                if not a["texto_util"]:
                    d["error"] = "PDF sin texto (escaneado): ábrelo a mano"
            except Exception as e:  # noqa: BLE001
                d["error"] = f"no se pudo leer el PDF ({type(e).__name__})"
        if analisis:
            f["requisitos"] = F.filas_requisitos(SimpleNamespace(analisis=analisis), raiz)
            f["notas"] = [x for x in f.get("notas", []) if not x.startswith("Pliegos sin analizar")]
    return len(candidatas), conseguidos


# ------------------------------------------------------------------ duplicados entre fuentes
def absorber_en_lic(destino, lic) -> None:
    if not any(d["url"] == lic.url for d in destino.duplicada_en):
        destino.duplicada_en.append({"fuente": lic.fuente, "url": lic.url})
    for k in ("presupuesto", "valor_estimado", "fecha_limite", "ccaa", "provincia", "ciudad", "pais"):
        if not getattr(destino, k) and getattr(lic, k):
            setattr(destino, k, getattr(lic, k))


def absorber_en_ficha(f: dict, lic) -> None:
    if not any(d["url"] == lic.url for d in f.setdefault("duplicada_en", [])):
        f["duplicada_en"].append({"fuente": lic.fuente, "url": lic.url})
    for k in ("presupuesto", "valor_estimado", "ccaa", "provincia", "ciudad", "pais"):
        if not f.get(k) and getattr(lic, k):
            f[k] = getattr(lic, k)
    if not f.get("fecha_limite") and lic.fecha_limite:
        f["fecha_limite"] = F.iso(lic.fecha_limite)


def main() -> int:
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:  # noqa: BLE001
        pass
    ap = argparse.ArgumentParser(description="Informe diario de licitaciones")
    ap.add_argument("--dias", type=int, help="días hacia atrás a revisar")
    ap.add_argument("--demo", action="store_true", help="ejemplo sin conexión")
    ap.add_argument("--sin-pliegos", action="store_true", help="no descargar ni analizar pliegos")
    ap.add_argument("--sin-ia", action="store_true", help="no usar el resumen con IA aunque haya clave")
    ap.add_argument("--fuentes", help="lista separada por comas: placsp,boe,ted,extra")
    ap.add_argument("--todo", action="store_true",
                    help="volver a evaluar también lo ya recogido (útil tras cambiar las palabras clave)")
    ap.add_argument("--abrir", action="store_true", help="abrir el Excel al terminar")
    ap.add_argument("--abrir-pdf", action="store_true", help="abrir el PDF al terminar")
    ap.add_argument("--abrir-web", action="store_true", help="abrir la web de los comerciales al terminar")
    ap.add_argument("--solo-web", action="store_true",
                    help="regenerar solo la web con lo ya recogido, sin consultar las fuentes")
    ap.add_argument("--reintentar-pliegos", action="store_true",
                    help="volver a descargar los pliegos que fallaron antes (licitaciones con plazo abierto) "
                         "y rehacer la web y el histórico, sin consultar las fuentes")
    args = ap.parse_args()

    hoy = date.today()
    dia = hoy.isoformat()
    raiz = Path(C.CARPETA_INFORMES)
    raiz = raiz if raiz.is_absolute() else BASE / raiz
    if args.demo:
        raiz = raiz / "DEMO"            # la demostración tiene su propia carpeta, con la misma estructura que la real
    carpeta = raiz / dia
    carpeta.mkdir(parents=True, exist_ok=True)

    print(f"Informe de licitaciones · {hoy:%d/%m/%Y}")
    try:
        reglas = palabras.cargar(BASE)
    except Exception as e:  # noqa: BLE001
        print(f"\nNo se pudo leer {C.ARCHIVO_PALABRAS}: {e}\nSi lo tienes abierto en Excel, ciérralo; si lo has roto, bórralo "
              "y se creará de nuevo con los valores por defecto.")
        return 1
    for aviso in reglas.avisos:
        print(f"  Aviso: {aviso}")

    if args.demo:
        ruta_reg = carpeta / "registro_demo.json"
        ruta_reg.unlink(missing_ok=True)
        registro = Registro(ruta_reg)
    else:
        registro = Registro(BASE / "registro.json")

    empresa_defecto = C.EMPRESA
    empresa = reglas.empresa or empresa_defecto
    C.EMPRESA = empresa                 # el PDF del día también usa este nombre
    logos = {"claro": reglas.logo_claro, "oscuro": reglas.logo_oscuro or reglas.logo_claro}

    if args.reintentar_pliegos:
        print("\nReintentando los pliegos que faltan…")
        revisadas, conseguidos = reintentar_pliegos(registro, raiz, hoy)
        print(f"\n{revisadas} licitaciones revisadas · {conseguidos} documentos descargados")
        if not args.demo:
            registro.guardar(hoy)
        informes.generar_historico(raiz / "HISTORICO_licitaciones.xlsx", registro.fichas(), raiz)

    if args.solo_web or args.reintentar_pliegos:
        try:
            ruta_web = exportar_web.exportar(raiz, registro.fichas(), registro.meta.get("ultimo_informe") or {},
                                             BASE, empresa, logos=logos, empresa_defecto=empresa_defecto)
        except Exception as e:  # noqa: BLE001
            print(f"No se pudo generar la web: {type(e).__name__}: {e}")
            return 1
        print(f"\nListo. Web: {ruta_web}")
        return 0

    if args.demo:
        from demo import datos_demo
        crudas = datos_demo(carpeta / "pliegos_demo")
        estados = [("DEMO", True, f"{len(crudas)} anuncios de ejemplo (sin conexión)")]
    else:
        desde = periodo(args, registro)
        activas = args.fuentes.split(",") if args.fuentes else C.FUENTES_ACTIVAS
        print(f"Periodo: desde {desde:%d/%m/%Y}")
        crudas, estados = recoger(desde, activas, registro)

    # ---- 1) cambios de estado de lo ya recogido (adjudicada, anulada...)
    cambios_estado, validas = [], []
    for lic in crudas:
        clave = f"{lic.fuente}|{lic.id}"
        if lic.estado and lic.estado not in C.ESTADOS_VALIDOS:
            f = registro.ficha(clave)
            if f and f.get("estado") != lic.estado:
                cambios_estado.append({"titulo": f["titulo"], "antes": f.get("estado"), "ahora": lic.estado})
                f["notas"] = f.get("notas", []) + [f"{hoy:%d/%m/%Y}: el estado pasa de {f.get('estado')} a {lic.estado}"]
                f["estado"] = lic.estado
            continue
        validas.append((clave, lic))

    # ---- 2) sin repetidos dentro de la lectura (nos quedamos con la versión más reciente)
    unicas: dict = {}
    for clave, lic in validas:
        previa = unicas.get(clave)
        if not previa or (lic.actualizada and (not previa.actualizada or lic.actualizada > previa.actualizada)):
            unicas[clave] = lic

    # ---- 3) ya conocidas / puntuación
    conocidas = descartadas = 0
    candidatas = []
    for clave, lic in unicas.items():
        marca = lic.actualizada.isoformat() if lic.actualizada else ""
        situacion = registro.situacion(clave, marca)
        if situacion == "igual" and not args.todo:
            conocidas += 1
            continue
        puntuar(lic, reglas)
        if lic.puntos < reglas.minima:
            registro.marcar_irrelevante(clave, marca, dia, lic.puntos)
            descartadas += 1
            continue
        candidatas.append((clave, marca, lic, situacion))

    # ---- 4) novedades, cambios y duplicados entre fuentes
    candidatas.sort(key=lambda x: (PRIORIDAD.get(x[2].fuente, 3), -x[2].puntos))
    en_curso: dict = {}          # clave -> lic de esta ejecución
    previas: dict = {}           # clave -> ficha de la ejecución anterior
    aceptadas: dict = {}         # clave -> ficha ligera, para comparar entre fuentes
    repetidas = 0
    for clave, marca, lic, situacion in candidatas:
        previa = registro.ficha(clave)
        primera = previa["primera"] if previa else dia
        ligera = F.ficha_de(lic, dia, primera, raiz, clave, marca)
        if previa:
            if situacion == "cambiada":
                cambios = F.comparar(previa, ligera)
                if not cambios:                       # actualización sin nada relevante: no se vuelve a listar
                    registro.entradas[clave]["u"] = marca
                    previa["marca"] = marca
                    conocidas += 1
                    continue
                lic.novedad, lic.cambios = "Actualizada", cambios
            else:
                lic.novedad = "Ya recogida"
            previas[clave] = previa
        else:
            lic.novedad = "Nueva"
            equivalente = registro.buscar_equivalente(lic, aceptadas)
            if equivalente:
                if equivalente in en_curso:
                    absorber_en_lic(en_curso[equivalente], lic)
                else:
                    absorber_en_ficha(registro.ficha(equivalente), lic)
                registro.marcar_irrelevante(clave, marca, dia, lic.puntos, duplicada_de=equivalente)
                repetidas += 1
                continue
        en_curso[clave] = lic
        aceptadas[clave] = ligera

    # ---- 5) completar BOE, descargar y analizar pliegos de lo nuevo o cambiado
    pendientes = [l for k, l in en_curso.items() if l.novedad in ("Nueva", "Actualizada")]
    print(f"\n{len(crudas)} leídas · {conocidas} ya conocidas · {repetidas} repetidas en otra fuente · "
          f"{descartadas} descartadas · {len(en_curso)} para el informe de hoy")
    if C.BOE_LEER_TEXTO and not args.demo:
        for lic in pendientes:
            if lic.fuente == "BOE":
                try:
                    fuentes.enriquecer_boe(lic)
                except Exception as e:  # noqa: BLE001
                    lic.notas.append(f"No se pudo abrir el texto del anuncio del BOE ({type(e).__name__}).")
    pendientes.sort(key=lambda l: -l.puntos)
    para_analizar = pendientes[:C.MAX_LICITACIONES_ANALIZADAS]
    for lic in pendientes[C.MAX_LICITACIONES_ANALIZADAS:]:
        lic.notas.append("Pliegos sin analizar: se alcanzó el límite diario de análisis.")
    if para_analizar:
        print("\nAnalizando pliegos…")
        analizar(para_analizar, carpeta / "pliegos", args.sin_pliegos, args.sin_ia)

    # ---- 6) fichas -> registro
    for clave, lic in en_curso.items():
        marca = lic.actualizada.isoformat() if lic.actualizada else ""
        previa = previas.get(clave)
        # (solo la demo usa «dias_atras_demo», para que la web de ejemplo tenga varios días)
        dia_f = (hoy - timedelta(days=getattr(lic, "dias_atras_demo", 0))).isoformat()
        ficha = F.ficha_de(lic, dia_f, previa["primera"] if previa else dia_f, raiz, clave, marca)
        if lic.novedad == "Ya recogida" and previa:      # no se vuelve a analizar: se conserva lo anterior
            for k in ("requisitos", "ia", "documentos", "notas", "duplicada_en"):
                ficha[k] = previa.get(k, ficha[k])
        registro.guardar_ficha(clave, marca, dia, ficha)

    # ---- 7) informes: Excel y PDF del día + histórico acumulado
    del_dia = [f for f in registro.fichas() if f.get("dia") == dia]
    del_dia.sort(key=lambda f: (-f["puntos"], f.get("fecha_limite") or "9999"))
    ctx = {"fecha": hoy, "fuentes": estados, "descartadas": descartadas, "repetidas": conocidas + repetidas,
           "demo": args.demo, "cambios_estado": cambios_estado}
    ruta_xlsx = carpeta / f"licitaciones_{dia}.xlsx"
    ruta_pdf = carpeta / f"licitaciones_{dia}.pdf"
    ruta_xlsx = informes.generar_excel_dia(ruta_xlsx, del_dia, ctx, raiz)
    informes.generar_pdf(ruta_pdf, del_dia, ctx, raiz)
    ruta_hist = informes.generar_historico(raiz / "HISTORICO_licitaciones.xlsx", registro.fichas(), raiz)

    info = {"dia": dia, "hora": datetime.now().strftime("%H:%M"), "fuentes": [list(e) for e in estados],
            "descartadas": descartadas, "conocidas": conocidas + repetidas}
    ruta_web = None
    if C.WEB_ACTIVA:
        try:
            ruta_web = exportar_web.exportar(raiz, registro.fichas(), info, BASE, empresa, logos=logos, empresa_defecto=empresa_defecto)
        except Exception as e:  # noqa: BLE001 — un fallo de la web no debe estropear el informe del día
            print(f"  ¡Atención! No se pudo actualizar la web: {type(e).__name__}: {e}")

    if not args.demo:
        registro.meta["ultimo_informe"] = info
        if all(ok for _, ok, _ in estados):    # si alguna fuente falla, mañana se revisa el mismo periodo
            registro.meta["ultima_ejecucion"] = datetime.now().isoformat(timespec="seconds")
        registro.guardar(hoy)

    print(f"\nListo. Carpeta del día: {carpeta}\n  Excel del día: {ruta_xlsx.name}\n  PDF del día:   {ruta_pdf.name}"
          f"\n  Histórico:     {ruta_hist}" + (f"\n  Web:           {ruta_web}" if ruta_web else ""))
    for lic in sorted(en_curso.values(), key=lambda l: -l.puntos)[:8]:
        print(f"   [{lic.puntos:>3}] {lic.novedad:<11} {lic.titulo[:90]}")
    for nombre_f, ok, texto in estados:
        if not ok:
            print(f"  ¡Atención! {nombre_f}: {texto}")
    if args.abrir or args.abrir_pdf or args.abrir_web:
        objetivo = ruta_web if (args.abrir_web and ruta_web) else (ruta_pdf if args.abrir_pdf else ruta_xlsx)
        try:
            if sys.platform.startswith("win"):
                os.startfile(objetivo)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{objetivo}"')
            else:
                os.system(f'xdg-open "{objetivo}" >/dev/null 2>&1 &')
        except Exception:  # noqa: BLE001
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
