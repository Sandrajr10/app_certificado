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
- Estructura prevista: `app.py` y `requirements.txt` en raíz; `templates/index.html`, `templates/resultado.html`, `templates/certificado.html`; `static/style.css`.
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
- App creada: `app.py`, `requirements.txt`, `iniciar_app.bat`, `templates/index.html`, `static/style.css` + `PLAN.md` y los dos `.xlsx` de `Insumos/`.
- Verificado 2026-10-15: Python 3.12 instalado vía winget + `pip install -r requirements.txt` (Flask, openpyxl). App corriendo en http://localhost:5000 con contadores 19/3/1 verificados.
- No hay `README`, manifiestos, tests, CI, ni `opencode.json`. No hay comandos de build/test/lint/typecheck; no inventarlos ni prometer su ejecución.
- Verificado 2026-10-14: no es repo git. No inicializar git ni reestructurar sin pedirlo.

## Insumos (solo lectura)
- No editar `Insumos/*.xlsx`. Para pruebas, pedir al usuario una copia o una exportación a CSV.
- `read` no extrae el contenido de los `.xlsx`; no asumir nada fuera de las columnas declaradas arriba sin pedir encabezados/ejemplo.
- Verificado 2026-10-14: `python`/`python3` son solo shims de Microsoft Store (sin intérprete real) y no hay `node`/`npm`; no intentar inspeccionar los Excel por `shell`.

## Entorno / Windows
- Shell = PowerShell. No usar encadenamiento estilo `||`/`&&` ni utilidades Unix (`cat`, `ls -la`, `find`); usar herramientas dedicadas `read`/`glob`/`grep`.
- La ruta contiene espacios y `OneDrive - Montecillos`: entrecomillar rutas y preferir el parámetro `workdir` en vez de `cd`.
- `read` sobre directorio ya muestra ocultos; no re-listar con `shell`.
