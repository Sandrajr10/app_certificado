# app.py — Academia Horizonte (LA COCINA)
# Toda la lógica vive aquí. La página solo muestra, no calcula.

import os  # para leer variables de entorno (configuración en la nube)
from datetime import date  # para la fecha de emisión del certificado
from pathlib import Path  # para armar rutas de archivos

from flask import Flask, render_template, request  # servidor web + páginas
from openpyxl import load_workbook  # para leer los Excel .xlsx

# VARIABLE de tipo objeto: la aplicación web.
# static_folder en public/: así Vercel sirve el CSS desde su CDN (recomendado por Vercel)
app = Flask(__name__, static_folder="public/static", static_url_path="/static")

# VARIABLES de tipo ruta: carpeta del proyecto y nombres fijos de los Excel
BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_MAESTRO = "Maestro_Estudiantes.xlsx"
ARCHIVO_REGISTRO = "Registro_Evaluaciones.xlsx"

# VARIABLE de tipo diccionario: caché simple para no releer los Excel en cada visita
CACHE = {"clave": None, "datos": None}


# FUNCIÓN: decide de qué carpeta se leen los Excel
def carpeta_insumos():
    # 1) Si existe la variable de entorno CARPETA_INSUMOS, se usa esa ruta
    elegida = os.environ.get("CARPETA_INSUMOS")
    if elegida:
        return Path(elegida)
    # 2) En el equipo local manda la carpeta real Insumos/
    local = BASE_DIR / "Insumos"
    if (local / ARCHIVO_MAESTRO).exists() and (local / ARCHIVO_REGISTRO).exists():
        return local
    # 3) En la nube (Vercel) se usan los datos ficticios de demo/
    return BASE_DIR / "demo"


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
    # VARIABLES de tipo ruta: ubicación de cada Excel según el entorno
    carpeta = carpeta_insumos()
    ruta_maestro = carpeta / ARCHIVO_MAESTRO
    ruta_registro = carpeta / ARCHIVO_REGISTRO

    # CONDICIONAL: si falta algún Excel se avisa en la página, no se rompe la app
    if not ruta_maestro.exists() or not ruta_registro.exists():
        aviso = (
            "No se encontraron los archivos de Insumos. "
            "Coloca Maestro_Estudiantes.xlsx y Registro_Evaluaciones.xlsx en Insumos/ "
            "o configura la variable de entorno CARPETA_INSUMOS."
        )
        return [], {"aprobacion": 0, "participacion": 0, "sin": 0}, [], [], aviso

    # CONDICIONAL: si los archivos no cambiaron desde la última lectura, se reutiliza
    clave = (
        str(ruta_maestro),
        ruta_maestro.stat().st_mtime_ns,
        str(ruta_registro),
        ruta_registro.stat().st_mtime_ns,
    )
    if CACHE["clave"] == clave:
        return CACHE["datos"]

    # --- 1. Leer el maestro: diccionario con clave (identificacion, programa) ---
    # VARIABLE de tipo diccionario: guarda los datos oficiales por estudiante+programa
    maestro = {}
    # VARIABLE de tipo conjunto: lista de programas sin repetidos (para el filtro)
    programas = set()
    libro = load_workbook(ruta_maestro, read_only=True, data_only=True)
    hoja = libro.active
    # BUCLE: recorre las filas del Excel, saltando el encabezado (fila 1)
    for fila in list(hoja.rows)[1:]:
        # CONDICIONAL: ignora filas vacías al final de la hoja
        if fila[0].value is None:
            continue
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
    libro2 = load_workbook(ruta_registro, read_only=True, data_only=True)
    hoja2 = libro2.active
    # BUCLE: recorre cada evaluación del registro
    for fila in list(hoja2.rows)[1:]:
        # CONDICIONAL: ignora filas vacías al final de la hoja
        if fila[0].value is None:
            continue
        # VARIABLES: identificación como texto, números como decimal
        ident = str(fila[0].value).strip()
        programa = str(fila[1].value).strip()
        nota = float(fila[3].value)
        asistencia = float(fila[4].value)
        clave_grupo = (ident, programa)
        # CONDICIONAL: si el grupo no existe, lo crea con listas vacías
        if clave_grupo not in grupos:
            grupos[clave_grupo] = {"notas": [], "asistencias": []}
        grupos[clave_grupo]["notas"].append(nota)
        grupos[clave_grupo]["asistencias"].append(asistencia)
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
                "cohorte": maestro[(ident, programa)]["cohorte"],
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

    # Se guarda el resultado en la caché antes de devolverlo
    CACHE["clave"] = clave
    CACHE["datos"] = (
        resultados,
        contadores,
        inconsistencias,
        sorted(programas),
        None,
    )
    return CACHE["datos"]


# Ruta principal: muestra la página con contadores, filtro, tabla e inconsistencias
@app.route("/")
def index():
    # VARIABLE de tipo texto: lo que el usuario eligió en el filtro (o "Todos")
    filtro = request.args.get("programa", "Todos")
    resultados, contadores, inconsistencias, programas, aviso = cargar_datos()
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
        aviso=aviso,
    )


# VARIABLE de tipo lista: nombres de los meses para la fecha del certificado
MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "setiembre", "octubre", "noviembre", "diciembre",
]


# FUNCIÓN: arma la fecha de hoy en texto (por ejemplo: 21 de octubre de 2026)
def fecha_emision():
    hoy = date.today()
    return str(hoy.day) + " de " + MESES[hoy.month - 1] + " de " + str(hoy.year)


# Ruta del certificado: se abre desde el enlace "Ver certificado" de la tabla
@app.route("/certificado/<identificacion>/<path:programa>")
def certificado(identificacion, programa):
    resultados, contadores, inconsistencias, programas, aviso = cargar_datos()
    # VARIABLE: aquí se guarda la fila que coincida con estudiante y programa
    encontrado = None
    # BUCLE: busca la combinación exacta (identificación + programa)
    for r in resultados:
        if r["identificacion"] == identificacion and r["programa"] == programa:
            encontrado = r
    # CONDICIONAL: si no existe, se muestra un mensaje en vez de un error feo
    if encontrado is None:
        return render_template(
            "mensaje.html",
            titulo="Certificado no encontrado",
            detalle="No existe un estudiante con esa identificación en ese programa.",
        ), 404
    # CONDICIONAL: regla de negocio: con asistencia menor a 80 no se emite certificado
    if encontrado["tipo"] == "Sin certificado":
        return render_template(
            "mensaje.html",
            titulo="No se emite certificado",
            detalle=encontrado["nombre"] + " tiene una asistencia de "
            + str(encontrado["asistencia"])
            + "%. El mínimo requerido es 80%.",
        )
    # Si cumple, se entrega el certificado ya calculado (el HTML no calcula nada)
    return render_template(
        "certificado.html",
        r=encontrado,
        fecha=fecha_emision(),
    )


# Arranque local: abre http://localhost:5000 en el navegador.
# En Vercel esta parte no se ejecuta: Vercel usa directamente el objeto "app".
if __name__ == "__main__":
    app.run(debug=True)
