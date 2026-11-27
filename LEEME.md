# Licitaciones diarias

Cada mañana revisa la Plataforma de Contratación del Sector Público, el BOE y TED (más los feeds que
añadas), puntúa lo que podría encajar con tu empresa, descarga los pliegos, localiza **en qué página**
está cada requisito y te deja un **Excel** (y un PDF) organizado por día. Recuerda lo que ya ha recogido,
así que no repite licitaciones. Funciona en tu ordenador: sin base de datos, sin Azure, sin servidores.

## 1. Instalación (una sola vez, en Windows)

Con Python instalado (python.org, marcando «Add python.exe to PATH»), abre una terminal en esta
carpeta y ejecuta:

    py -m pip install -r requirements.txt

## 2. Primera prueba, sin conexión

    py licitaciones.py --demo

Crea `informes\DEMO\` con un Excel, un PDF y un histórico de ejemplo (datos inventados) para que veas
el aspecto.

## 3. Uso diario

    py licitaciones.py --abrir

| Qué | Dónde |
|---|---|
| Excel del día (lo nuevo o cambiado hoy) | `informes\AAAA-MM-DD\licitaciones_AAAA-MM-DD.xlsx` |
| PDF del día, con enlaces a la página de cada requisito | `informes\AAAA-MM-DD\licitaciones_AAAA-MM-DD.pdf` |
| Pliegos descargados | `informes\AAAA-MM-DD\pliegos\` |
| **Histórico de todo lo recogido, sin duplicados** | `informes\HISTORICO_licitaciones.xlsx` |
| Memoria de lo ya visto (no la toques) | `registro.json` |

**Columnas del Excel:** encaje y puntos; objeto y órgano; novedad y qué ha cambiado; **palabras clave
detectadas y dónde** (título o lote); comunidad autónoma, provincia, ciudad y país; **presupuesto total**
y **el importe de lo que te interesa** (si hay lotes, suma solo los lotes de interés; la hoja *Lotes* los
detalla); valor estimado; fecha límite de presentación y días restantes (se pone en rojo a 5 días o
menos); tipo de contrato, procedimiento, duración, CPV; por qué encaja; en qué páginas del pliego está cada
requisito (solvencia, garantías, ISO, criterios…); enlaces al anuncio, al pliego descargado y al original;
si está también publicada en otra fuente; avisos.

El histórico incluye *Fecha de búsqueda* y *Primera recogida*, y una hoja «Resumen por día».
Si tienes el Excel abierto al ejecutar, el programa guarda una copia con «_nuevo» en el nombre.

**Si fallan los pliegos.** Cada pliego se intenta descargar hasta 3 veces (con 4 y 8 segundos de pausa) si
el servidor da error o devuelve una página web en lugar del PDF. Si aun así falla, la ficha dice el motivo
real. Para volver a intentar más tarde los que falten (solo licitaciones con plazo abierto, las de más
puntos primero), sin volver a consultar las fuentes:

    py licitaciones.py --reintentar-pliegos

Actualiza las fichas de la web y el histórico. El Excel y el PDF de aquel día se quedan como estaban.

## 4. Dónde cambiar las palabras clave

Abre **`palabras_clave.xlsx`** (se crea solo la primera vez). Guarda, **ciérralo** y ya se aplica en la
siguiente ejecución. Hojas:

- **Grupos:** familias de servicio y los puntos que da cada una.
- **Palabras:** qué se busca. Añade filas, cámbialas o pon NO en «Activa» para desactivar una sin borrarla.
  «Normal» acepta terminaciones (escribe la raíz: `museo` encuentra museos, museográfico); «Exacta» es
  solo para siglas (VR, CMS, app). Da igual mayúsculas y tildes.
- **CPV:** prefijos de códigos y sus puntos. **Exclusiones:** palabras que restan (limpieza, obras…).
  Se aplican también a cada lote.
- **Ajustes:** puntuación mínima para aparecer, umbral de encaje ALTO, importe mínimo y máximo. Además,
  tres filas de texto: **Nombre de la empresa**, **Logo para el tema claro** y **Logo para el tema
  oscuro**. En el panel de la web se editan en la pestaña **Empresa**.

**Nombre y logo de la web.** En *Configuración > Empresa* escribes el nombre y eliges un logo para cada
tema entre las imágenes de la carpeta `web\config` del programa. Para añadir otro logo, copia la imagen
(PNG, JPG o SVG, mejor cuadrada) en esa carpeta y ejecuta el programa: aparecerá en la lista. Al guardar,
el cambio se ve al momento en tu navegador. Para todos, se ve a partir de la siguiente ejecución, o al
momento con `py licitaciones.py --solo-web`. Si el nombre se deja vacío, se usa `EMPRESA` de `configuracion.py`.

Si cambias palabras y quieres revisar de nuevo lo de los últimos días:

    py licitaciones.py --dias 7 --todo

Lo ya recogido saldrá como «Ya recogida» y el histórico se actualiza sin duplicar.

## 5. Cómo evita duplicados

- Cada licitación se identifica por su fuente e identificador. Si ya estaba y no ha cambiado en origen,
  no vuelve a salir.
- Si cambia (plazo, importe, estado, documentos nuevos), sale como **«Actualizada»** con lo que ha cambiado.
- Si el mismo contrato aparece en dos fuentes (por ejemplo PLACSP y BOE), se reconoce por el expediente o
  por el objeto y el órgano, y se muestra una sola vez con «También publicada en».
- Lo descartado por baja relevancia se recuerda 120 días para no reevaluarlo cada día.

## 6. Que se ejecute solo cada mañana

El ordenador tiene que estar encendido. Programador de tareas de Windows → Crear tarea básica → Diaria
8:00 → Iniciar un programa: `py`, argumentos `licitaciones.py`, «Iniciar en»: la ruta de esta carpeta.

## 7. La web para los comerciales (OneDrive)

Cada ejecución deja también una web estática dentro de la carpeta de informes: `index.html` + `datos\`.
No necesita servidor, usuarios ni internet. Inicio con mapa por comunidades, novedades y plazos que
cierran; listado con filtros y CSV; ficha con presupuesto, lotes, palabras clave y página de cada requisito;
vista por día con enlaces al Excel y PDF.

1. En `configuracion.py` pon `CARPETA_INFORMES` en una carpeta de OneDrive, p. ej.
   `r"C:\Users\valen\OneDrive - Mi empresa\Licitaciones"`, y compártela (solo lectura) con los comerciales.
2. Programa `ejecutar_diario.bat` con el Programador de tareas (diaria, 8:00). Registra en `ejecucion.log`.
3. Los comerciales abren `index.html` **desde su carpeta sincronizada** (clic derecho → «Mantener siempre en
   este dispositivo»). No uses la vista web de OneDrive: puede no ejecutar la página.

Prueba sin conexión: `py licitaciones.py --demo --abrir-web`. Para rehacer la web sin buscar:
`py licitaciones.py --solo-web`. La web es de solo lectura: no guarda quién se queda cada oferta.

**Compartir una licitación por correo.** En la ficha hay dos botones. *Enviar correo* abre Outlook con
el mensaje ya maquetado (título, importes, plazo, por qué encaja y un botón al anuncio original), con
el *Para* vacío para que pongas tú al destinatario. No adjunta archivos: solo enlaces. *Copiar con
formato* copia ese mismo bloque al portapapeles, manteniendo negritas y enlaces, para pegarlo con
`Ctrl+V` en un correo que ya tengas abierto.

Para que *Enviar correo* abra Outlook de un solo clic hay que registrar una vez el protocolo
`licitaciones:`. Son dos pasos, y solo la primera vez:

    py preparar_protocolo.py

y después doble clic en **`registrar_protocolo.reg`** (es para tu usuario, no pide administrador).
La primera vez Chrome pregunta si permite abrir la aplicación: marca «Permitir siempre». Quien no lo
tenga registrado —por ejemplo los comerciales, que no tienen el programa— se descarga un `.eml` que
al abrirlo hace lo mismo. Para desactivarlo, `quitar_protocolo.reg`.

Los `.reg` se generan en vez de venir hechos porque llevan dentro la ruta de la carpeta y la del
Python instalado: si mueves el programa o cambias de versión de Python, vuelve a ejecutar
`preparar_protocolo.py` y a importar el archivo. Quien atiende la llamada es `compartir_correo.py`,
que lee la ficha de `informes\datos\f\` y se la pasa a Outlook. Con el paquete **pywin32** instalado
(`py -m pip install pywin32`) el correo además conserva tu firma; sin él sale sin ella.

Los pulgares *Sí, me interesa* y *No me interesa* se ven en la ficha y en el correo, pero todavía no
registran nada: de momento son solo visuales.

## 8. Otras opciones

- `--dias 3` revisa los últimos 3 días · `--sin-pliegos` va más rápido · `--fuentes placsp,boe`.
- Otras fuentes con RSS/Atom (boletines autonómicos, universidades…): `FEEDS_EXTRA` en `configuracion.py`.
- Resumen redactado con IA (opcional): define la variable de entorno `ANTHROPIC_API_KEY`. Sin ella, el
  informe usa extractos literales del pliego con su página.

## Límites que conviene conocer

- **Las conexiones a PLACSP, BOE y TED están escritas según su documentación pública y probadas con
  muestras, no contra los servidores reales.** En los primeros días puede que alguna fuente necesite un
  ajuste: el informe la marca en rojo en la hoja *Fuentes* y el resto sigue funcionando.
- En los anuncios del BOE, plazo, importes y lugar se leen del texto con búsqueda aproximada: verifícalos.
- La comunidad autónoma sale del código oficial cuando existe; si no, se deduce del texto y se marca «(aprox.)».
  Los ministerios y organismos estatales pueden quedar sin comunidad.
- Un PDF escaneado no se puede leer; algunas webs solo dejan descargar el pliego desde el navegador. El
  informe lo avisa y enlaza al original.
- Si un lote no indica importe, el «importe de lo que te interesa» queda incompleto y se avisa.
- El informe ayuda a decidir, no sustituye la lectura del pliego oficial.
