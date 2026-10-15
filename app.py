# app.py — Academia Horizonte (LA COCINA)
# Toda la lógica vive aquí. La página solo muestra, no calcula.

from pathlib import Path  # para armar rutas de archivos

from flask import Flask, render_template, request  # servidor web + páginas
from openpyxl import load_workbook  # para leer los Excel .xlsx

# VARIABLE de tipo objeto: la aplicación web
app = Flask(__name__)

# VARIABLES de tipo ruta: Bodega en solo lectura, ubicada desde este archivo
# pathlib + Path(__file__): funciona sin importar desde qué carpeta se arranque la app
BASE_DIR = Path(__file__).resolve().parent
RUTA_MAESTRO = BASE_DIR / "Insumos" / "Maestro_Estudiantes.xlsx"
RUTA_REGISTRO = BASE_DIR / "Insumos" / "Registro_Evaluaciones.xlsx"


# FUNCIÓN: calcula el promedio simple de una lista de números
def calcular_promedio(valores):
    # VARIABLE de tipo número: acumulador que empieza en 0
    total = 0
    # BUCLE: recorre cada valor de la lista, uno por uno
    for v in valores:
        total = total + v
    # devuelve el total dividido entre la cantidad (promedio simple)
    return total / len(valores)


# FUNCIÓN: decide el tipo de certificado según las reglas (límites INCLUYEN el valor)
def decidir_certificado(promedio, asistencia):
    # CONDICIONAL 1: sin certificado si la asistencia es menor de 80
    if asistencia < 80:
        return "Sin certificado"
    # CONDICIONAL 2: aprobación si el promedio es 70 o más (y asistencia ya >= 80)
    elif promedio >= 70:
        return "Aprobación"
    # CONDICIONAL 3: lo demás con buena asistencia es participación
    else:
        return "Participación"


# FUNCIÓN: lee los Excel y prepara todo lo que la página necesita
def cargar_datos():
    # --- 1. Leer el maestro: diccionario con clave (identificacion, programa) ---
    # VARIABLE de tipo diccionario: guarda los datos oficiales por estudiante+programa
    maestro = {}
    # VARIABLE de tipo conjunto: lista de programas sin repetidos (para el filtro)
    programas = set()
    libro = load_workbook(RUTA_MAESTRO, read_only=True, data_only=True)
    hoja = libro.active
    # BUCLE: recorre las filas del Excel, saltando el encabezado (fila 1)
    for fila in list(hoja.rows)[1:]:
        # VARIABLES de tipo texto: la identificación SIEMPRE como texto en el cruce
        ident = str(fila[0].value).strip()
        nombre = str(fila[1].value).strip()
        programa = str(fila[3].value).strip()
        maestro[(ident, programa)] = {
            "nombre": nombre,
            "correo": str(fila[2].value).strip(),
            "cohorte": str(fila[4].value).strip(),
            "programa": programa,
        }
        programas.add(programa)
    libro.close()

    # --- 2. Leer el registro: agrupar notas y asistencias por (identificacion, programa) ---
    # VARIABLE de tipo diccionario: cada clave guarda sus listas de notas y asistencias
    grupos = {}
    libro2 = load_workbook(RUTA_REGISTRO, read_only=True, data_only=True)
    hoja2 = libro2.active
    # BUCLE: recorre cada evaluación del registro
    for fila in list(hoja2.rows)[1:]:
        # VARIABLES: identificación como texto, números como decimal
        ident = str(fila[0].value).strip()
        programa = str(fila[1].value).strip()
        nota = float(fila[3].value)
        asistencia = float(fila[4].value)
        clave = (ident, programa)
        # CONDICIONAL: si el grupo no existe, lo crea con listas vacías
        if clave not in grupos:
            grupos[clave] = {"notas": [], "asistencias": []}
        grupos[clave]["notas"].append(nota)
        grupos[clave]["asistencias"].append(asistencia)
    libro2.close()

    # --- 3. Calcular promedios y clasificar (solo grupos que están en el maestro) ---
    # VARIABLE de tipo lista: filas finales de la tabla de resultados
    resultados = []
    # VARIABLE de tipo lista: evaluaciones sin maestro (huérfanos, no emitibles)
    huerfanos = []
    # BUCLE: un certificado por estudiante y programa
    for (ident, programa), datos in grupos.items():
        promedio = calcular_promedio(datos["notas"])
        asistencia = calcular_promedio(datos["asistencias"])
        tipo = decidir_certificado(promedio, asistencia)
        # CONDICIONAL: sin ficha en el maestro no se puede emitir
        if (ident, programa) not in maestro:
            huerfanos.append({
                "identificacion": ident,
                "programa": programa,
                "modulos": len(datos["notas"]),
                "promedio": round(promedio, 2),
                "asistencia": round(asistencia, 2),
                "tipo": tipo,
            })
        else:
            resultados.append({
                "identificacion": ident,
                "nombre": maestro[(ident, programa)]["nombre"],
                "programa": programa,
                "modulos": len(datos["notas"]),
                "promedio": round(promedio, 2),
                "asistencia": round(asistencia, 2),
                "tipo": tipo,
            })

    # --- 4. Maestro sin ninguna evaluación ---
    # VARIABLE de tipo lista: estudiantes oficiales que no aparecen en el registro
    sin_evaluacion = []
    # BUCLE: revisa cada estudiante oficial
    for (ident, programa), datos in maestro.items():
        # CONDICIONAL: si su clave no está en los grupos, no tiene notas
        if (ident, programa) not in grupos:
            sin_evaluacion.append({
                "identificacion": ident,
                "nombre": datos["nombre"],
                "programa": programa,
            })

    # --- 5. Ordenar la tabla por programa y nombre (solo presentación) ---
    resultados.sort(key=lambda r: (r["programa"], r["nombre"]))

    # --- 6. Contadores de la parte de arriba (Cocina los calcula, Salón los muestra) ---
    # VARIABLES de tipo número: empiezan en 0
    cont_aprobacion = 0
    cont_participacion = 0
    cont_sin = 0
    # BUCLE: cuenta cada fila según su tipo
    for r in resultados:
        # CONDICIONALES: suma 1 al contador que corresponda
        if r["tipo"] == "Aprobación":
            cont_aprobacion = cont_aprobacion + 1
        elif r["tipo"] == "Participación":
            cont_participacion = cont_participacion + 1
        else:
            cont_sin = cont_sin + 1

    # --- 7. Sección INCONSISTENCIAS (las 3 detectadas en la validación) ---
    # VARIABLE de tipo lista: cada elemento describe un caso raro
    inconsistencias = []
    # BUCLE: cada estudiante oficial sin notas es una inconsistencia
    for s in sin_evaluacion:
        inconsistencias.append({
            "caso": "Maestro sin evaluaciones",
            "detalle": s["nombre"] + " (" + s["identificacion"] + ") — "
            + s["programa"] + " no tiene ninguna nota registrada.",
        })
    # BUCLE: cada evaluación sin maestro es otra inconsistencia
    for h in huerfanos:
        inconsistencias.append({
            "caso": "Evaluación sin maestro",
            "detalle": "Identificación " + h["identificacion"] + " — "
            + h["programa"] + " tiene " + str(h["modulos"])
            + " módulos pero no existe en el maestro (no se puede emitir).",
        })
    # BUCLE: cada fila sin certificado (ej. abandono con 0) se reporta aparte
    for r in resultados:
        # CONDICIONAL: solo los "Sin certificado"
        if r["tipo"] == "Sin certificado":
            inconsistencias.append({
                "caso": "Sin certificado por asistencia",
                "detalle": r["nombre"] + " (" + r["identificacion"] + ") — "
                + r["programa"] + ": promedio " + str(r["promedio"])
                + ", asistencia " + str(r["asistencia"]) + ".",
            })

    contadores = {
        "aprobacion": cont_aprobacion,
        "participacion": cont_participacion,
        "sin": cont_sin,
    }
    return resultados, contadores, inconsistencias, sorted(programas)


# Ruta principal: muestra la página con contadores, filtro, tabla e inconsistencias
@app.route("/")
def index():
    # VARIABLE de tipo texto: lo que el usuario eligió en el filtro (o "Todos")
    filtro = request.args.get("programa", "Todos")
    resultados, contadores, inconsistencias, programas = cargar_datos()
    # CONDICIONAL: si hay filtro, deja solo ese programa; si no, muestra todo
    if filtro != "Todos":
        # BUCLE con CONDICIONAL implícito: filtra la lista por programa
        resultados = [r for r in resultados if r["programa"] == filtro]
    # entrega todo ya calculado a la plantilla (el HTML no calcula nada)
    return render_template(
        "index.html",
        resultados=resultados,
        contadores=contadores,
        inconsistencias=inconsistencias,
        programas=programas,
        filtro=filtro,
    )


# Arranque: abre http://localhost:5000 en el navegador
if __name__ == "__main__":
    app.run(debug=True)
