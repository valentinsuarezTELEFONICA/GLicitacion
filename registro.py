# -*- coding: utf-8 -*-
"""
Registro de todo lo ya recogido (registro.json): evita duplicados, detecta cambios y alimenta el histórico.

Cada licitación se identifica por «fuente|id». Si es relevante se guarda su ficha completa;
si no lo es, solo una marca para no volver a evaluarla mientras no cambie en origen.
"""
from __future__ import annotations

import json
import re
from datetime import date, timedelta
from difflib import SequenceMatcher
from pathlib import Path

from textos import norm

VERSION = 2
DIAS_IRRELEVANTES = 120      # cuánto se recuerdan las descartadas


def _alnum(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", norm(s))


def _parecido(a: str, b: str) -> float:
    a, b = norm(a), norm(b)
    return SequenceMatcher(None, a, b).ratio() if a and b else 0.0


class Registro:
    def __init__(self, ruta: Path):
        self.ruta = Path(ruta)
        self.meta: dict = {"ultima_ejecucion": None}
        self.entradas: dict = {}
        self._cargar()

    # ---------------------------------------------------------------- persistencia
    def _cargar(self) -> None:
        try:
            datos = json.loads(self.ruta.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            datos = None
        if datos and datos.get("version") == VERSION:
            self.meta = datos.get("meta", self.meta)
            self.entradas = datos.get("entradas", {})
            return
        # migración desde la versión anterior (estado.json)
        viejo = self.ruta.with_name("estado.json")
        try:
            v1 = json.loads(viejo.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        self.meta["ultima_ejecucion"] = v1.get("ultima_ejecucion")
        for clave, v in v1.get("vistas", {}).items():
            self.entradas[clave] = {"u": v.get("u", ""), "visto": v.get("visto", ""), "relevante": False, "puntos": 0}

    def guardar(self, hoy: date | None = None) -> None:
        hoy = hoy or date.today()
        corte = (hoy - timedelta(days=DIAS_IRRELEVANTES)).isoformat()
        self.entradas = {k: v for k, v in self.entradas.items()
                         if v.get("relevante") or v.get("visto", "9999") >= corte}
        tmp = self.ruta.with_suffix(".tmp")
        tmp.write_text(json.dumps({"version": VERSION, "meta": self.meta, "entradas": self.entradas},
                                  ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.replace(self.ruta)

    # ---------------------------------------------------------------- consulta
    def situacion(self, clave: str, marca: str) -> str:
        """'nueva' (nunca vista), 'igual' (ya vista sin cambios en origen) o 'cambiada'."""
        e = self.entradas.get(clave)
        if not e:
            return "nueva"
        return "igual" if e.get("u") == marca else "cambiada"

    def ficha(self, clave: str) -> dict | None:
        e = self.entradas.get(clave)
        return e.get("ficha") if e else None

    def fichas(self) -> list:
        return [e["ficha"] for e in self.entradas.values() if e.get("ficha")]

    # ---------------------------------------------------------------- escritura
    def marcar_irrelevante(self, clave: str, marca: str, hoy: str, puntos: int, duplicada_de: str = "") -> None:
        previa = self.entradas.get(clave, {})
        e = {"u": marca, "visto": hoy, "relevante": False, "puntos": puntos}
        if duplicada_de:
            e["duplicada_de"] = duplicada_de
        if previa.get("ficha"):             # sigue en el histórico aunque ahora puntúe poco
            e["ficha"], e["relevante"] = previa["ficha"], True
        self.entradas[clave] = e

    def guardar_ficha(self, clave: str, marca: str, hoy: str, ficha: dict) -> None:
        self.entradas[clave] = {"u": marca, "visto": hoy, "relevante": True,
                                "puntos": ficha.get("puntos", 0), "ficha": ficha}

    # ---------------------------------------------------------------- duplicados entre fuentes
    def buscar_equivalente(self, lic, aceptadas: dict | None = None) -> str | None:
        """Clave de una ficha de OTRA fuente que parece ser la misma licitación (mismo expediente u objeto)."""
        candidatas = {k: e["ficha"] for k, e in self.entradas.items() if e.get("ficha")}
        candidatas.update(aceptadas or {})
        exp = _alnum(lic.expediente)
        for clave, f in candidatas.items():
            if f.get("fuente") == lic.fuente:
                continue
            org = _parecido(lic.organo, f.get("organo", ""))
            tit = _parecido(lic.titulo, f.get("titulo", ""))
            if len(exp) >= 5 and exp == _alnum(f.get("expediente", "")) and (org > 0.45 or tit > 0.5):
                return clave
            if tit > 0.92 and org > 0.6:
                return clave
        return None
