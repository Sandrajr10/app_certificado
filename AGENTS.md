# AGENTS.md

## Contexto
- Academia Horizonte, centro de formación técnica para empresas. Al cerrar cada cohorte emite un certificado por estudiante. Proceso actual manual en Excel/Word; objetivo: automatizarlo como aplicación web.

## Arquitectura prevista (aún no existe código)
- Bodega: `Insumos/*.xlsx` (nótese mayúscula). Solo lectura, nunca modificar.
- Cocina (backend): `app.py` con Flask en la raíz. Toda la lógica y cálculos van aquí, nunca en la interfaz.
- Salón (frontend): HTML servido por Flask. Solo muestra; no calcula ni decide.

## Decisiones técnicas (ver PLAN.md para el orden)
- Excel con `openpyxl` (lee `.xlsx` directo, liviana). Sin `pandas` (oculta la lógica) y sin `xlrd` (ya no lee `.xlsx`).
- Web con `Flask` (mínimo, `app.py` + `templates/`). Sin `FastAPI`/`Django` (sobran para esto).
- Estructura actual: `app.py`, `requirements.txt`, `vercel.json` y `.vercelignore` en raíz; `templates/login.html`, `templates/_barra_usuario.html`, `templates/index.html`, `templates/inconsistencias.html`, `templates/sin_permiso.html`, `templates/certificado.html`, `templates/mensaje.html`; `public/static/style.css`, `public/static/certificado.css`; `demo/*.xlsx` (datos ficticios para la nube); `.env.ejemplo` (sí se sube) y `.env` (ignorado).
- `Identificacion` siempre como texto en el cruce.

## Entradas (columnas verificadas contra el binario el 2026-10-15 vía XML interno)
- `Insumos/Maestro_Estudiantes.xlsx`: Identificacion, Nombre_Completo, Correo, Programa, Cohorte.
- `Insumos/Registro_Evaluaciones.xlsx`: Identificacion, Programa, Modulo, Nota (0-100), Asistencia_Pct, Fecha_Cierre.

## Reglas de negocio (límites inclusivos, no cambiar sin pedir)
- Agrupar por Identificacion + Programa; un certificado por estudiante y programa.
- Promedio = suma de Notas / cantidad de módulos cursados. Asistencia = promedio de Asistencia_Pct.
- Aprobación si Promedio >= 70 y Asistencia >= 80. Participación si Promedio < 70 y Asistencia >= 80. Sin certificado si Asistencia < 80.

## Estilo de trabajo
- Código Python simple y comentado en español, marcando variables, tipos, condicionales, bucles y funciones.
- Diseño web azul marino y dorado.
- Explicar en español sencillo lo que se vaya haciendo.

## Estado actual
- App creada: `app.py`, `requirements.txt`, `iniciar_app.bat`, `templates/index.html`, `public/static/style.css` + `PLAN.md`, `README.md` y los dos `.xlsx` de `Insumos/`.
- Verificado 2026-10-15: Python 3.12 instalado vía winget + `pip install -r requirements.txt` (Flask, openpyxl). App corriendo en http://localhost:5000 con contadores 19/3/1 verificados.
- Verificado 2026-10-21: panel de administrador con carga en memoria y descarga Excel. Pruebas locales: ver subidos (19 aprobaciones), descargar subidos (24 filas con encabezado), descargar fábrica, archivo faltante, extensión incorrecta, Excel inválido, envío de 5 MB (413) y usuario normal (403, sin panel).
- Repo git activo en rama `main` con remoto `origin` = https://github.com/Sandrajr10/app_certificado.git. `Insumos/` queda fuera del remoto (`.gitignore` + `.vercelignore`): datos personales. No reestructurar sin pedirlo.

## Despliegue (Vercel)
- Vercel detecta Flask automáticamente: `app.py` expone `app`, las dependencias están en `requirements.txt` y no se define comando de build. `vercel.json` solo ajusta `maxDuration`.
- En la nube no existe `Insumos/`; `carpeta_insumos()` cae en `demo/` (Excel ficticios, no personas reales).
- La variable de entorno `CARPETA_INSUMOS` permite apuntar a otra carpeta de datos sin tocar código.
- Los estáticos van en `public/` (Vercel los sirve por CDN); Flask usa `static_folder="public/static"` para que también funcione en local.
- `cargar_datos()` tiene caché por fecha de modificación de los Excel y devuelve un aviso si faltan archivos.

## Rutas
- `/login` — pantalla de entrada; valida correo y contraseña contra variables de entorno.
- `/salir` — borra la sesión y vuelve al login.
- `/` — tabla con contadores, filtro por programa e inconsistencias (estas solo para rol `admin`). Requiere sesión.
- `/inconsistencias` — página completa de inconsistencias, exclusiva de `admin`; con rol distinto responde 403 y `templates/sin_permiso.html`. Requiere sesión.
- `/certificado/<identificacion>/<programa>` — certificado imprimible (Aprobación o Participación). Requiere sesión. Si la asistencia es < 80 muestra el motivo; si no existe la combinación, 404 con mensaje.
- El certificado incluye cohorte, promedio, asistencia y fecha de emisión; el HTML no calcula nada.
- `/cargar` — POST exclusivo de `admin` (`solo_admin`). Recibe `maestro` y `registro` con `enctype="multipart/form-data"` y el campo `accion` (`ver` o `descargar`). Procesa los Excel en memoria con `io.BytesIO`; nunca los guarda. Con `ver` muestra el tablero calculado en esa misma respuesta; con `descargar` devuelve un Excel armado en memoria con `send_file`.
- `/descargar` — GET exclusivo de `admin`. Descarga en memoria los resultados de los datos de fábrica (enlace del panel).

## Acceso y sesiones
- Flask `session` (cookie firmada). Sin base de datos todavía.
- Credenciales en variables de entorno: `ADMIN_CORREO`/`ADMIN_CLAVE` (rol `admin`) y `USUARIO_CORREO`/`USUARIO_CLAVE` (rol `usuario`). Nunca en el código.
- Si falta cualquiera de las cuatro, se activa `MODO_DEMO` con credenciales de prueba visibles en `/login`.
- `SECRET_KEY` firma la sesión; en Vercel debe configurarse como variable de entorno.
- `requiere_login` protege `/`, `/certificado/...`, `/inconsistencias` y `/salir`; redirige a `/login` si no hay sesión.
- `solo_admin` exige rol `admin` y responde 403 con `sin_permiso.html`; está aplicado a `/inconsistencias`.
- En `/`, la lista de inconsistencias se vacía para roles distintos de `admin` y la sección se oculta en la plantilla. El rol se muestra como insignia (`rol-admin` dorada, `rol-usuario` clara).
- `python-dotenv` carga `.env` en local; `.env` está en `.gitignore` y `.vercelignore`.

## Carga de Excel (panel admin)
- El panel solo se renderiza si `rol_sesion == "admin"`; la protección real está en `/cargar` con `solo_admin`.
- `cargar_datos(archivo_maestro=None, archivo_registro=None)` acepta dos archivos en memoria; si no se los pasa, lee los Excel de fábrica. Los datos subidos no usan la caché (cada archivo es distinto); los de fábrica sí.
- `crear_excel(resultados)` arma el `.xlsx` en un `io.BytesIO` y se envía con `send_file`; no se escribe en disco.
- Límite: `app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024` (4 MB). El manejador 413 vuelve al tablero con aviso rojo, porque Vercel corta cerca de 4.5 MB.
- Validaciones con aviso rojo y sin página en blanco: faltan archivos, extensión distinta de `.xlsx`, Excel inválido/columnas incorrectas o tamaño excesivo. El detalle técnico se registra con `app.logger.warning`.
- El panel muestra siempre el origen: "Datos de fábrica" o los nombres subidos. Con datos subidos aparece un aviso verde: viven solo en esa pantalla y al recargar se vuelve a fábrica (no hay almacenamiento).
- Los resultados subidos no se guardan, así que los enlaces de certificado se ocultan ("No disponible") y el enlace a `/inconsistencias` no aparece en esa pantalla; el Excel descargado sí contiene todos los resultados.

## Insumos (solo lectura)
- No editar `Insumos/*.xlsx`. Para pruebas, pedir al usuario una copia o una exportación a CSV.
- `read` no extrae el contenido de los `.xlsx`; no asumir nada fuera de las columnas declaradas arriba sin pedir encabezados/ejemplo.
- Verificado 2026-10-14: `python`/`python3` son solo shims de Microsoft Store (sin intérprete real) y no hay `node`/`npm`; no intentar inspeccionar los Excel por `shell`.

## Entorno / Windows
- Shell = PowerShell. No usar encadenamiento estilo `||`/`&&` ni utilidades Unix (`cat`, `ls -la`, `find`); usar herramientas dedicadas `read`/`glob`/`grep`.
- La ruta contiene espacios y `OneDrive - Montecillos`: entrecomillar rutas y preferir el parámetro `workdir` en vez de `cd`.
- `read` sobre directorio ya muestra ocultos; no re-listar con `shell`.
