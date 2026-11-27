"""Escribe los dos .reg que dan de alta y de baja el protocolo «licitaciones:».

Se generan en lugar de guardarlos hechos porque dentro llevan la ruta de esta
carpeta y la del Python instalado, que cambian de un equipo a otro. Ejecuta

    py preparar_protocolo.py

y luego doble clic en registrar_protocolo.reg. Haz lo mismo si mueves la carpeta
del programa o cambias de versión de Python.
"""
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
BARRA = chr(92)


def dobles(ruta):
    """En un .reg las barras de las rutas van dobladas."""
    return str(ruta).replace(BARRA, BARRA * 2)


def entrecomillado(texto):
    return BARRA + '"' + texto + BARRA + '"'


def clave(*partes):
    return BARRA.join(("HKEY_CURRENT_USER", "Software", "Classes") + partes)


def main():
    lanzador = Path(sys.executable).with_name("pythonw.exe")
    if not lanzador.exists():
        lanzador = Path(sys.executable)
    guion = BASE / "compartir_correo.py"
    if not guion.exists():
        print("No encuentro compartir_correo.py en", BASE)
        return 1

    orden = " ".join([
        entrecomillado(dobles(lanzador)),
        entrecomillado(dobles(guion)),
        entrecomillado("%1"),
    ])

    alta = [
        "Windows Registry Editor Version 5.00",
        "",
        "; Da de alta el protocolo licitaciones: para este usuario; no hace falta",
        "; ser administrador. Con el, al pulsar «Enviar correo» en la web, Chrome",
        "; llama a compartir_correo.py y el correo se abre maquetado en Outlook.",
        "; La primera vez Chrome pregunta si permite abrir la aplicacion: marca",
        "; «Permitir siempre». Para quitarlo, ejecuta quitar_protocolo.reg.",
        "; Generado por preparar_protocolo.py: no lo edites a mano.",
        "",
        "[" + clave("licitaciones") + "]",
        '@="URL:Agente Oportunidades"',
        '"URL Protocol"=""',
        "",
        "[" + clave("licitaciones", "shell", "open", "command") + "]",
        '@="' + orden + '"',
        "",
    ]
    baja = [
        "Windows Registry Editor Version 5.00",
        "",
        "; Quita el protocolo licitaciones: de este usuario. El boton «Enviar",
        "; correo» de la web vuelve a descargar un .eml en lugar de abrir Outlook.",
        "",
        "[-" + clave("licitaciones") + "]",
        "",
    ]

    for nombre, lineas in (("registrar_protocolo.reg", alta), ("quitar_protocolo.reg", baja)):
        (BASE / nombre).write_text("\r\n".join(lineas), encoding="utf-16")
        print("escrito:", nombre)
    print()
    print("Ahora haz doble clic en registrar_protocolo.reg.")
    print("Lanzara:", lanzador)
    return 0


if __name__ == "__main__":
    sys.exit(main())
