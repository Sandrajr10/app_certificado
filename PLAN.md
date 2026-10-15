# PLAN.md — Academia Horizonte (app de certificados)

Modo Plan: este archivo manda. No escribir código hasta que el usuario apruebe cada paso.

## Orden de construcción

1. Validar el cruce de datos (sin código de app, solo lectura).
2. Backend con Flask (`app.py`, Cocina).
3. Interfaz HTML (Salón, solo muestra).

## Paso 0 — Requisito previo (una vez)

- El entorno actual no tiene Python real (solo shims de Microsoft Store). El usuario instala Python 3 + pip fuera de la sesión.
- Luego se usarán `Flask` + `openpyxl` (decisión en `AGENTS.md`). Sin pandas, sin `xlrd`.
- Validación: el usuario confirma `python --version` y que puede instalar paquetes. Sin esto no se empieza el Paso 1.

## Paso 1 — Validar el cruce de datos

Qué haré:
- Leer `Insumos/*.xlsx` como solo lectura, sin modificarlos.
- Confirmar conteos base: 24 en maestro, 88 filas en registro, columnas esperadas.
- Clasificar cada grupo `Identificacion + Programa` según reglas de `AGENTS.md`.
- Reportar las 3 listas: emitibles, huérfanos (notas sin maestro) y sin notas / sin derecho.

Casos testigo ya detectados (no asumir, verificar en el reporte):
- `304560321`: en maestro, sin notas → queda fuera, pedir al usuario cómo mostrarlo.
- `999880777`: con notas, sin maestro → no emitible, va a lista de revisión.
- `111230984`: `44.5 / 36.25` → sin certificado.

Validación antes de seguir:
- El usuario revisa el reporte y decide: ¿el huérfano bloquea o solo se lista como "revisar"? ¿Qué texto mostramos para "sin notas" y "sin certificado"?
- Esperado si nada cambió: 22 emitibles (19 Aprobación + 3 Participación), 1 sin derecho, 1 huérfano, 1 sin notas. Si los números difieren, se detiene y se investiga.

## Paso 2 — Backend Flask (`app.py`)

Qué haré (cuando el usuario lo pida):
- Crear `app.py` en la raíz + `requirements.txt` (`Flask`, `openpyxl`).
- Toda la lógica vive en `app.py`: carga de Excel, agrupación por `Identificacion + Programa`, promedios, decisión Aprobación / Participación / Sin (límites inclusivos `>=70`, `>=80`).
- Rutas mínimas previstas: `/` (buscador), `/resultado` (detalle por estudiante+programa), `/certificado` (vista imprimible). Nombres exactos se confirman al implementar.
- `Identificacion` se trata como texto en todo el cruce.

Validación antes de seguir:
- Levantar el servidor local y probar 4 casos: un Aprobación claro (`305680912`), un Participación (`116890234`), un Sin (`111230984`) y un inexistente/huérfano (`999880777`).
- El usuario confirma que los 4 resultados coinciden con el reporte del Paso 1. Sin ese visto bueno no se toca la interfaz.

## Paso 3 — Interfaz HTML (Salón)

Qué haré (cuando el usuario lo pida):
- `templates/index.html`, `templates/resultado.html`, `templates/certificado.html` + `static/style.css` en azul marino y dorado.
- Los HTML solo muestran lo que `app.py` calcula. Cero lógica en plantillas.
- Código Python simple y comentado en español, marcando variables, tipos, condicionales, bucles y funciones.

Validación final:
- Recorrido visual con el usuario: buscar, ver resultado, imprimir certificado.
- Comprobar: no se modificó ningún `.xlsx`, no hay cálculo en el HTML, explicación en español sencillo.
- Cierre: el usuario acepta o pide ajustes por escrito antes de dar por terminada cada plantilla.

## Regla de avance

- Un paso a la vez, con confirmación explícita del usuario entre pasos.
- Si los datos cambian (nuevos Excel), se repite el Paso 1 antes de tocar código.
