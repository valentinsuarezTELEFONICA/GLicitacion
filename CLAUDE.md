# Agente Oportunidades: contexto del proyecto

Documento para retomar el desarrollo en otra sesión de Claude. Estado a 4/11/2026.

## 1. Qué es y para quién
Herramienta interna para la empresa de Valen (VR/AR, experiencias inmersivas, museos y centros de interpretación,
hardware, audiovisual y videowalls, software a medida, gestores documentales y CMS, apps, gamificación, educación,
cultura y turismo). Cada mañana lee fuentes de licitaciones públicas, puntúa lo que encaja, descarga los pliegos,
localiza en qué página está cada requisito y genera Excel + PDF por día. Los **comerciales** consumen el resultado en
una **web estática de solo lectura** alojada en una carpeta compartida de OneDrive.

Restricciones de Valen: "para andar por casa": sin Azure, sin base de datos, sin servidores, sin usuarios ni roles,
todo interno y sin salir a internet en la parte web (los permisos los da OneDrive). Windows, comando `py`.

## 2. Arquitectura
Python 3 (requests, openpyxl, reportlab, pymupdf). Tarea programada diaria → `licitaciones.py` → lee fuentes →
puntúa → descarga y analiza pliegos → escribe Excel/PDF/histórico → `exportar_web.py` genera `index.html` + `datos/`.

**Carpeta buena: `C:\Users\rt00484.IOTBD\Documents\Claude\Licitaciones`** (el programa está en la raíz).
`Desktop\licitaciones_diarias\v1`, `v3` y `v4` son copias antiguas; v4 tiene el mismo código Python que había
aquí antes del 9/11/2026, además de los borradores `index_v0..v4.html`. Por ahora solo se usa en local; lo de
login, roles y producción vendrá más adelante, cuando lo diga Valen.

Carpeta del programa:
- `licitaciones.py`: orquestador. Flags: `--demo`, `--dias N`, `--todo`, `--sin-pliegos`, `--fuentes`, `--abrir`,
  `--abrir-pdf`, `--abrir-web`, `--solo-web` (rehace la web sin buscar), `--reintentar-pliegos` (desde el
  10/11/2026: vuelve a bajar los pliegos que faltan en fichas con plazo abierto, rehace requisitos, histórico
  y web; no regenera el Excel ni el PDF de aquel día).
- `fuentes.descargar` (corregido el 10/11/2026): hasta 3 intentos ante 5xx/429/página HTML, cierra el archivo
  antes de borrarlo y reintenta si el antivirus lo bloquea (en este equipo hay Cortex XDR y Defender).
  Antes, cualquier respuesta que no fuera un PDF acababa como `WinError 32`, porque se borraba el archivo con
  el handle abierto. Por eso fallaron todos los pliegos del 4/11 y del 9/11.
- Empresa y logo (10/11/2026): filas de texto de la hoja Ajustes ("Nombre de la empresa", "Logo para el tema
  claro/oscuro"), editables en la pestaña Empresa del panel. `palabras.Reglas.empresa/logo_claro/logo_oscuro`
  → `exportar_web` publica `D.empresa`, `D.logo`, `D.logos` (imágenes de `web/config`, copiadas a
  `informes/config`). `pintarMarca()` en `index.html` pinta la cabecera, y el panel la llama al guardar.
- `fuentes.py`: PLACSP (feed Atom/CODICE paginado, lotes, códigos NUTS), BOE (API datos abiertos, sección 5),
  TED (API v3, campos opcionales probados uno a uno), feeds RSS/Atom extra (`FEEDS_EXTRA`).
- `analisis.py` + `palabras.py`: puntuación por palabras clave, CPV, exclusiones; puntuación por lote
  (`minima_lote`=20); `importe_interes` = suma de lotes de interés; exclusiones también a nivel de lote.
- `registro.py` (`registro.json`): deduplicación por "fuente|id", detección de "Actualizada" (compara fichas),
  equivalencia entre fuentes (expediente igual + similitud órgano/título → "También publicada en"), descartadas
  recordadas 120 días, migración del antiguo `estado.json`.
- `fichas.py`: modelo común (dict) para Excel, PDF y web. `ubicacion.py`: CCAA/provincia/ciudad (dato, NUTS o
  inferido "aprox."). `informes.py`: Excel diario, PDF e histórico. `textos.py`, `demo.py` (20 licitaciones de
  ejemplo en varias CCAA y días, con `dias_atras_demo`).
- `exportar_web.py`: escribe `datos/indice.js` (`window.DATOS`, lista ligera), `datos/f/<sha1[:12]>.js` (ficha
  completa, carga perezosa) y copia `web/index.html` a la raíz. Escritura atómica y solo si cambia (amigable con
  OneDrive). Se usan `.js` y no `.json` porque `file://` bloquea `fetch`.
- `web/index.html`: el diseño que usa Valen (desde el 9/11/2026 es una copia exacta de `informes/index.html`,
  de 143 KB): paleta Telefónica, mapa SVG de España con burbujas y zoom, estadísticas por grupo, router por hash.
  `web/config/` (`configuracion.js` y `jszip.min.js`) contiene el panel de configuración. `exportar_web` copia todo
  byte a byte y solo si cambia. **Los cambios de diseño se hacen en `web/`**: si se toca solo
  `informes/index.html`, la siguiente ejecución lo pisa. `web/index_anterior_sin_mapa.html` es la plantilla
  vieja (mapa de teselas), que se guarda solo como referencia.
- `configuracion.py`: fuentes, ventana de días, `CARPETA_INFORMES` (ruta de OneDrive), `WEB_ACTIVA`, conceptos a
  buscar en pliegos, IA opcional (`ANTHROPIC_API_KEY`, modelo `claude-sonnet-5-5`).
- `palabras_clave.xlsx`: hojas Instrucciones, Grupos, Palabras (Normal = raíz + terminaciones / Exacta), CPV,
  Exclusiones, Ajustes. Es la **única** fuente de reglas. Se edita con Excel o con el panel Configuración de
  la web, que escribe en el mismo archivo.
- Compartir por correo (27/11/2026): en la ficha, `Enviar correo` y `Copiar con formato`, más los
  pulgares de interés (solo visuales, sin acción todavía). El cuerpo se arma en HTML con tablas y
  estilos en línea porque Outlook lo pinta con el motor de Word; `mailto:` quedó descartado porque
  solo admite texto plano, y `ActiveXObject` porque murió con Internet Explorer. El botón navega a
  `licitaciones:compartir?id=<id corto de la ficha>`, protocolo registrado en
  `HKCU\Software\Classes\licitaciones`. Los dos .reg los escribe `preparar_protocolo.py` y están en
  el `.gitignore`, porque llevan dentro la ruta del programa y la de Python; si se mueve la carpeta
  hay que regenerarlos. Las funciones del correo se exponen en `window` (`window.enviarCorreo`),
  porque todo el script del index vive dentro de `(function(){…})()` con "use strict" y los onclick
  se evalúan en el ámbito global: sin eso el botón no hace nada.
  `compartir_correo.py` lee `informes/datos/f/<id>.js`, valida el id contra `[0-9a-zA-Z_-]{6,64}`,
  escapa todos los campos y abre Outlook por COM si hay pywin32 (así conserva la firma) o, si no,
  escribiendo un `.eml` con `X-Unsent: 1` en la carpeta temporal y abriéndolo. Si el protocolo no
  está registrado, la web espera 1,5 s y descarga el `.eml` ella misma. La plantilla del correo está
  duplicada a propósito en `web/index.html` y en `compartir_correo.py`: sin servidor no hay forma de
  compartirla. `prueba_compartir_mail.html` es la maqueta de referencia, con vista previa al lado.
- `ejecutar_diario.bat`: `py licitaciones.py >> ejecucion.log`, para el Programador de tareas.

Estructura de salida: `informes/AAAA-MM-DD/{licitaciones_*.xlsx, .pdf, pliegos/}`, `informes/HISTORICO_licitaciones.xlsx`,
`informes/index.html`, `informes/datos/`. Demo en `informes/DEMO/`.

Claves de una ficha: clave, marca, id, fuente, dia, primera, novedad (Nueva/Actualizada/Ya recogida), cambios,
puntos, encaje (ALTO/MEDIO), titulo, expediente, organo, descripcion, tipo, procedimiento, duracion, pais, ccaa,
provincia, ciudad, ubicacion_origen, presupuesto, valor_estimado, n_lotes, n_lotes_interes, importe_interes,
fecha_limite, cpv, palabras [{grupo,termino,donde}], lotes, requisitos [{concepto,paginas,ruta_rel,extracto}],
documentos, ia, notas, duplicada_en. `ruta_rel` es relativa a la carpeta de informes.

## 3. Estado actual (lo construido en esta sesión)
- Programa completo y probado con simulación (dedup, cambios, cruce PLACSP↔BOE, edición de palabras).
- Web v1 probada en Chromium bajo `file://` sin errores: Inicio (mapa de teselas por CCAA con "+N hoy", cierran
  en 7 días, Hoy, Novedades), Todas (filtros en la URL, orden, CSV), Ficha (presupuesto, lotes, palabras clave,
  pliegos, requisito → página con `#page=N`), Por día, aviso de fuentes caídas o datos de más de 48 h, modo
  oscuro, responsive.
- Entregados: `licitaciones_diarias.zip` y `web_demo.zip`.

## 4. Peticiones posteriores de Valen
Estado revisado el 9/11/2026:
Las cuatro ya estaban hechas en el diseño de Valen (`informes/index.html`, del 4/11/2026). El diagnóstico del
9/11 se equivocó porque revisó la plantilla vieja.
1. **Pantalla de configuración**: `config/configuracion.js` lee y guarda **directamente `palabras_clave.xlsx`**
   con la File System Access API (Edge/Chrome), con descarga como alternativa. Tiene deshacer, validación y
   borrado de grupo eligiendo qué hacer con sus palabras.
   El 9/11/2026 se probó un mecanismo paralelo con `config.json` y **se retiró ese mismo día** por indicación de
   Valen: duplicaba el panel. No hay que volver a proponerlo.
2. **Mapa de España** en SVG, con burbujas por comunidad, zoom y enlace al listado (`montarMapa`).
3. **Paleta Telefónica**: `--marca:#0066ff` y secundarios `--g1..g8`.
4. **Orden del Inicio**: `estadisticas()` arriba; debajo, a la izquierda el mapa y "Cierran", a la derecha
   "Hoy" y "Novedades".

## 5. Limitaciones y riesgos conocidos
- PLACSP, BOE y TED están escritos según la documentación pública y probados con muestras, no contra los
  servidores reales (el sandbox no tenía acceso). Los primeros días pueden requerir ajustes; el informe marca en
  rojo la fuente que falle.
- En el BOE, plazo, importe y lugar se extraen del texto con búsqueda aproximada.
- La web es solo lectura: no guarda quién se queda cada oferta (haría falta lista SharePoint o servidor).
- La vista web de OneDrive puede no ejecutar el HTML: abrirlo desde la carpeta sincronizada ("Mantener siempre en
  este dispositivo"). Sin verificar.
- La tarea programada solo corre con el PC encendido y OneDrive sincronizado. Equipo por decidir.
- La tipografía final (Bahnschrift/Segoe UI) no se vio en las capturas (el sandbox solo tenía DejaVu).
- PDFs escaneados no se pueden leer; algunas webs solo permiten descargar el pliego desde el navegador.

## 6. Cómo arrancar
    py -m pip install -r requirements.txt
    py licitaciones.py --demo --abrir-web      # prueba sin conexión
    py licitaciones.py --abrir-web             # uso real
En `configuracion.py` fijar `CARPETA_INFORMES = r"C:\Users\...\OneDrive - ...\Licitaciones"` y programar
`ejecutar_diario.bat` a las 8:00.

## 7. Siguientes pasos sugeridos
1. Primera ejecución real y corrección de fuentes según los errores de la hoja *Fuentes*.
2. Hecho en el diseño de Valen: panel de configuración sobre el Excel, paleta Telefónica, mapa con burbujas,
   orden del Inicio y estadísticas por grupo (ver sección 4).
3. Revisar en Chrome la web real (`informes/index.html`, no la de DEMO) después de la ejecución del 9/11/2026.
4. Elegir dónde corre la tarea programada (hoy no hay ninguna creada).
