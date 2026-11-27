"""Abre en Outlook un correo con la licitación que ha pedido la web.

Lo llama Windows cuando en la web se pulsa «Enviar correo», que navega a
licitaciones:compartir?id=<id>. El id es el nombre del archivo de la ficha
dentro de informes/datos/f/.

Por qué existe: un enlace mailto: solo admite texto plano, así que no hay forma
de que el correo salga con negritas, colores ni enlaces. Registrando este
protocolo, el clic sale del navegador y llega aquí, donde sí se puede hablar
con Outlook.

Dos maneras de abrirlo, de mejor a peor:
  1. Con pywin32 (COM): Outlook crea el mensaje, le deja su firma y encima se
     le pone el cuerpo maquetado.
  2. Sin pywin32: se escribe un .eml en la carpeta temporal y se abre. Outlook
     lo presenta como mensaje nuevo gracias a la cabecera X-Unsent, pero sin
     la firma.

Se registra con registrar_protocolo.reg y se quita con quitar_protocolo.reg.
"""
import html
import json
import os
import re
import sys
import tempfile
import urllib.parse
from base64 import b64encode
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

ID_VALIDO = re.compile(r"^[0-9a-zA-Z_-]{6,64}$")
TIPO = "font-family:Segoe UI,Arial,sans-serif"


def aviso(texto, titulo="Agente Oportunidades"):
    """Sin consola (pythonw) la única forma de contar un fallo es un cuadro."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, texto, titulo, 0x40)
    except Exception:
        print(texto)


def carpeta_informes():
    import configuracion as C
    raiz = Path(C.CARPETA_INFORMES)
    return raiz if raiz.is_absolute() else (BASE / raiz)


def id_pedido(args):
    """Saca el id de licitaciones:compartir?id=xxx (o del primer argumento)."""
    if not args:
        return None
    crudo = args[0]
    partido = urllib.parse.urlsplit(crudo)
    consulta = partido.query or (crudo.split("?", 1)[1] if "?" in crudo else "")
    valores = urllib.parse.parse_qs(consulta)
    pedido = (valores.get("id") or valores.get("clave") or [""])[0].strip()
    return pedido or None


def leer_ficha(id_ficha):
    """Lee informes/datos/f/<id>.js, que es JSON envuelto en una asignación."""
    archivo = carpeta_informes() / "datos" / "f" / (id_ficha + ".js")
    if not archivo.exists():
        return None
    texto = archivo.read_text(encoding="utf-8")
    if "]=" not in texto:
        return None
    return json.loads(texto.split("]=", 1)[1].rstrip().rstrip(";"))


def e(valor):
    """Algunos campos de la ficha, como «donde», vienen en lista."""
    if isinstance(valor, (list, tuple)):
        valor = ", ".join(str(x) for x in valor if x)
    return html.escape("" if valor is None else str(valor), quote=True)


def dinero(valor):
    if valor is None or valor == "":
        return ""
    try:
        return format(float(valor), ",.0f").replace(",", ".") + " €"
    except (TypeError, ValueError):
        return str(valor)


def fecha(valor):
    if not valor:
        return ""
    texto = str(valor).replace("T", " ")
    try:
        dia = texto.split(" ")[0]
        hora = texto.split(" ")[1][:5]
        a, m, d = dia.split("-")
        return d + "/" + m + "/" + a + " " + hora
    except (IndexError, ValueError):
        return texto


def celda(titulo, valor, azul=False):
    if not valor:
        return ""
    tam = "13pt" if azul else "11pt"
    color = "#0066ff" if azul else "#111827"
    return (
        '<td width="33%" valign="top" style="padding:0 6px 12px 0;">'
        '<table width="100%" cellpadding="0" cellspacing="0"'
        ' style="border:1px solid #e3e8ef;background:#f7f9fc;">'
        '<tr><td style="padding:10px 12px;">'
        '<div style="' + TIPO + ';font-size:8.5pt;color:#6b7280;text-transform:uppercase;'
        'letter-spacing:.4px;padding-bottom:3px;">' + e(titulo) + '</div>'
        '<div style="' + TIPO + ';font-size:' + tam + ';font-weight:bold;color:' + color + ';">'
        + e(valor) + '</div>'
        '</td></tr></table></td>'
    )


def etiqueta(texto, fondo, tinta):
    return (
        '<td bgcolor="' + fondo + '" style="padding:4px 10px;' + TIPO
        + ';font-size:8.5pt;font-weight:bold;color:' + tinta + ';">'
        + e(texto) + '</td><td width="8"></td>'
    )


def lugar(f):
    vistos = set()
    limpio = []
    for p in (f.get("ciudad"), f.get("provincia"), f.get("ccaa")):
        p = (p or "").strip()
        if p and p.lower() not in vistos:
            vistos.add(p.lower())
            limpio.append(p)
    return ", ".join(limpio)


def correo_html(f):
    etiquetas = ""
    if f.get("novedad") == "Nueva":
        etiquetas += etiqueta("Nueva", "#f1e0fa", "#471551")
    elif f.get("novedad") == "Actualizada":
        etiquetas += etiqueta("Actualizada", "#c5f8f9", "#253d3c")
    if f.get("encaje") == "ALTO":
        etiquetas += etiqueta("Encaje alto", "#fce3cd", "#63442a")
    elif f.get("encaje") == "MEDIO":
        etiquetas += etiqueta("Encaje medio", "#f8f0bd", "#473515")

    datos = [("Presupuesto total", dinero(f.get("presupuesto")), True)]
    if f.get("importe_interes") is not None and (f.get("n_lotes") or 0) > 1:
        datos.append((
            "Lo que te interesa (" + str(f.get("n_lotes_interes")) + " de "
            + str(f.get("n_lotes")) + " lotes)",
            dinero(f.get("importe_interes")), True))
    datos += [
        ("Valor estimado", dinero(f.get("valor_estimado")), False),
        ("Fecha límite", fecha(f.get("fecha_limite")), False),
        ("Publicada", fecha(f.get("fecha_publicacion")), False),
        ("Tipo y procedimiento",
         " · ".join(x for x in (f.get("tipo"), f.get("procedimiento")) if x), False),
        ("Duración", f.get("duracion"), False),
        ("Puntos de encaje", f.get("puntos"), False),
    ]
    datos = [d for d in datos if d[1] not in (None, "", "—")]

    filas = ""
    for n in range(0, len(datos), 3):
        tres = datos[n:n + 3]
        filas += "<tr>" + "".join(celda(*d) for d in tres)
        if len(tres) < 3:
            filas += '<td width="' + str(33 * (3 - len(tres))) + '%"></td>'
        filas += "</tr>"

    borde = "padding:5px 10px;border-bottom:1px solid #eef1f6;" + TIPO + ";font-size:10.5pt;"
    palabras = ""
    for p in f.get("palabras") or []:
        palabras += (
            "<tr>"
            '<td style="' + borde + '"><strong>' + e(p.get("termino")) + "</strong></td>"
            '<td style="' + borde + 'color:#374151;">' + e(p.get("grupo")) + "</td>"
            '<td style="' + borde + 'color:#6b7280;">' + e(p.get("donde")) + "</td>"
            "</tr>"
        )

    trozos = [f.get("organo"), lugar(f), f.get("fuente")]
    if f.get("expediente"):
        trozos.append("Expte. " + str(f["expediente"]))
    cabecera = " &middot; ".join(e(x) for x in trozos if x)

    url = (f.get("url") or "").strip()
    if not url.lower().startswith(("http://", "https://")):
        url = ""

    cab_sec = ('<div style="' + TIPO + ';font-size:9pt;font-weight:bold;color:#6b7280;'
               'text-transform:uppercase;letter-spacing:.6px;border-top:1px solid #e3e8ef;'
               'padding:16px 0 8px;">')
    col = 'padding:6px 10px;' + TIPO + ';font-size:9pt;color:#6b7280;'

    h = [
        '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>',
        '<body style="margin:0;padding:0;background:#ffffff;">',
        '<table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;">',
        '<tr><td style="padding:4px 2px 18px;">',
        '<div style="' + TIPO + ';font-size:10pt;color:#6b7280;padding-bottom:10px;">'
        'Licitación de interés</div>',
    ]
    if etiquetas:
        h.append('<table cellpadding="0" cellspacing="0"><tr>' + etiquetas + "</tr></table>")
    h += [
        '<div style="' + TIPO + ';font-size:15pt;font-weight:bold;color:#0066ff;'
        'line-height:1.3;padding:14px 0 6px;">' + e(f.get("titulo")) + '</div>',
        '<div style="' + TIPO + ';font-size:10pt;color:#6b7280;padding-bottom:18px;">'
        + cabecera + '</div>',
        '<table width="100%" cellpadding="0" cellspacing="0">' + filas + '</table>',
    ]
    if palabras:
        h += [
            cab_sec + "Por qué encaja</div>",
            '<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e3e8ef;">',
            '<tr bgcolor="#f7f9fc">',
            '<td style="' + col + '">Palabra clave</td>',
            '<td style="' + col + '">Familia</td>',
            '<td style="' + col + '">Dónde aparece</td>',
            "</tr>", palabras, "</table>",
        ]
    if f.get("motivos"):
        h.append('<div style="' + TIPO + ';font-size:10pt;color:#374151;padding:10px 0 0;">'
                 + " &middot; ".join(e(m) for m in f["motivos"]) + "</div>")
    if url:
        h += [
            '<div style="' + TIPO + ';font-size:10.5pt;color:#111827;padding:20px 0 10px;">'
            "Consulta el anuncio completo en:</div>",
            '<table cellpadding="0" cellspacing="0"><tr>'
            '<td bgcolor="#0066ff" style="padding:11px 22px;">',
            '<a href="' + e(url) + '" style="' + TIPO + ';font-size:10.5pt;font-weight:bold;'
            'color:#ffffff;text-decoration:none;">ABRIR EL ANUNCIO ORIGINAL</a>',
            "</td></tr></table>",
        ]
    h += [
        '<div style="' + TIPO + ';font-size:10.5pt;color:#111827;'
        'border-top:1px solid #e3e8ef;margin-top:26px;padding:16px 0 10px;">'
        "&iquest;Esta licitaci&oacute;n crees que te resulta de inter&eacute;s?</div>",
        '<table cellpadding="0" cellspacing="0"><tr>',
        '<td bgcolor="#dcfce7" style="padding:9px 18px;border:1px solid #86efac;">',
        '<span style="font-family:Segoe UI Emoji,Arial,sans-serif;font-size:13pt;">&#128077;</span>',
        '<span style="' + TIPO + ';font-size:10pt;font-weight:bold;color:#166534;">'
        "&nbsp;S&iacute;, me interesa</span></td>",
        '<td width="10"></td>',
        '<td bgcolor="#fee2e2" style="padding:9px 18px;border:1px solid #fca5a5;">',
        '<span style="font-family:Segoe UI Emoji,Arial,sans-serif;font-size:13pt;">&#128078;</span>',
        '<span style="' + TIPO + ';font-size:10pt;font-weight:bold;color:#b91c1c;">'
        "&nbsp;No me interesa</span></td>",
        "</tr></table>",
        '<div style="' + TIPO + ';font-size:8.5pt;color:#9aa6bd;'
        'border-top:1px solid #e3e8ef;margin-top:22px;padding-top:10px;">'
        "Compartido desde el Buscador de Licitaciones</div>",
        "</td></tr></table></body></html>",
    ]
    return "".join(h)


def por_com(asunto, cuerpo):
    """Outlook pone su firma al abrir el mensaje; el cuerpo va encima de ella."""
    import win32com.client as com
    mensaje = com.Dispatch("Outlook.Application").CreateItem(0)
    mensaje.Subject = asunto
    mensaje.Display(False)
    mensaje.HTMLBody = cuerpo + (mensaje.HTMLBody or "")
    return "com"


def por_eml(asunto, cuerpo):
    """X-Unsent: 1 es lo que hace que Outlook lo abra como mensaje nuevo."""
    texto = b64encode(cuerpo.encode("utf-8")).decode("ascii")
    partido = "\r\n".join(texto[i:i + 76] for i in range(0, len(texto), 76))
    eml = "\r\n".join([
        "To: ",
        "Subject: =?UTF-8?B?" + b64encode(asunto.encode("utf-8")).decode("ascii") + "?=",
        "X-Unsent: 1",
        "MIME-Version: 1.0",
        "Content-Type: text/html; charset=UTF-8",
        "Content-Transfer-Encoding: base64",
        "", partido, "",
    ])
    destino = Path(tempfile.gettempdir()) / "licitacion_compartida.eml"
    destino.write_text(eml, encoding="utf-8", newline="")
    os.startfile(str(destino))
    return "eml"


def main():
    pedido = id_pedido(sys.argv[1:])
    if not pedido:
        aviso("No se ha indicado qué licitación compartir.")
        return 2
    if not ID_VALIDO.match(pedido):
        aviso("El identificador de la licitación no es válido.")
        return 2

    try:
        f = leer_ficha(pedido)
    except Exception as err:
        aviso("No se ha podido leer la ficha.\n\n" + str(err))
        return 1
    if not f:
        aviso("No se encuentra la ficha de esta licitación.\n\n"
              "Ejecuta el programa para regenerar la web:\n"
              "py licitaciones.py --solo-web")
        return 1

    asunto = "Buscador de Licitaciones: " + (f.get("titulo") or "")
    cuerpo = correo_html(f)
    try:
        por_com(asunto, cuerpo)
    except Exception:
        try:
            por_eml(asunto, cuerpo)
        except Exception as err:
            aviso("No se ha podido abrir el correo en Outlook.\n\n" + str(err))
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
