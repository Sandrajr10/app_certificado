# app_certificado — Academia Horizonte

Aplicación web en Flask que lee dos Excel, calcula el promedio y la asistencia de
cada estudiante y decide qué certificado le corresponde (Aprobación, Participación
o Sin certificado).

## Estructura

```
app.py                      # Toda la lógica (la "cocina")
requirements.txt            # Flask + openpyxl + python-dotenv
iniciar_app.bat             # Arranque local con doble clic
templates/login.html        # Pantalla de entrada (correo y contraseña)
templates/_barra_usuario.html  # Correo, rol y botón Salir
templates/index.html        # Página (el "salón": solo muestra)
templates/inconsistencias.html  # Página completa, solo administradores
templates/sin_permiso.html  # Respuesta amable del error 403
public/static/style.css     # Estilos (Vercel los sirve desde public/)
demo/                       # Excel ficticios para la demo pública
Insumos/                    # Excel reales (NO se suben: datos personales)
.env.ejemplo                # Nombres de las variables (sí se sube)
.env                        # Credenciales locales (NO se sube)
vercel.json                 # Configuración de despliegue
.vercelignore               # Evita subir Insumos/, .env y archivos innecesarios
```

## Ejecutar en local

1. Instalar dependencias:

   ```powershell
   python -m pip install -r requirements.txt
   ```

2. Arrancar la app (o hacer doble clic en `iniciar_app.bat`):

   ```powershell
   python app.py
   ```

3. Abrir <http://localhost:5000>

La app usa `Insumos/*.xlsx` si existen; si no, usa `demo/*.xlsx`.

## Acceso

La app pide correo y contraseña antes de mostrar cualquier dato. Las cuentas se
leen de variables de entorno; nunca están escritas en el código.

1. Copiar `.env.ejemplo` como `.env` y completar los valores:

   ```
   SECRET_KEY=una-clave-larga-y-aleatoria
   ADMIN_CORREO=...
   ADMIN_CLAVE=...
   USUARIO_CORREO=...
   USUARIO_CLAVE=...
   ```

2. Si falta alguna de las cuatro variables de cuenta, la app entra en **modo
   demostración** y muestra en la pantalla de entrada las credenciales de prueba.

3. En Vercel, configurar esas mismas variables en *Settings → Environment
   Variables*. `SECRET_KEY` es obligatoria para que la sesión se mantenga entre
   instancias.

Roles actuales: `admin` y `usuario`.

- **Administrador:** ve los contadores, la tabla y la sección de inconsistencias,
  además de la página completa `/inconsistencias`.
- **Usuario:** ve solo los contadores y la tabla de certificados. Si intenta abrir
  `/inconsistencias`, recibe un error 403 con una página explicativa.

La protección está en `app.py` con el decorador `solo_admin`; el HTML solo oculta
la sección para no confundir, pero no es la barrera de seguridad.

## Despliegue en Vercel

El repositorio ya está listo para Vercel (detección automática de Flask, sin
comando de build):

1. Entrar a <https://vercel.com/new> e importar el repositorio
   `Sandrajr10/app_certificado`.
2. No cambiar la configuración: Vercel detecta `app.py` y `requirements.txt`.
3. Pulsar **Deploy**. La app queda publicada en una URL `*.vercel.app`.

Cada `git push` a `main` genera un despliegue nuevo.

### Datos usados en la nube

Vercel **no** recibe la carpeta `Insumos/` (está en `.gitignore` y en
`.vercelignore`). En su lugar se publican los datos ficticios de `demo/`, creados
solo para demostrar el funcionamiento: no corresponden a personas reales.

Para usar otros datos sin cambiar el código, se puede definir la variable de
entorno `CARPETA_INSUMOS` con la ruta de la carpeta que contiene los dos Excel.
Por ejemplo, en un servidor propio:

```powershell
$env:CARPETA_INSUMOS = "D:\datos\cohorte-2026"
python app.py
```

> Importante: no subir datos personales reales a un repositorio público. La app
> muestra nombres e identificaciones en la página.

## Certificados

En la tabla de resultados, cada estudiante con derecho a certificado tiene un
enlace **Ver certificado**. La vista está lista para imprimir o guardar como PDF
con el botón **Imprimir / Guardar PDF** (tamaño A4 horizontal).

- **Aprobación:** promedio ≥ 70 y asistencia ≥ 80.
- **Participación:** promedio < 70 y asistencia ≥ 80.
- **Sin certificado:** la vista indica el motivo (asistencia menor a 80).

## Reglas de negocio

- Se agrupa por `Identificacion + Programa`: un certificado por estudiante y programa.
- Promedio = suma de las notas / cantidad de módulos cursados.
- Asistencia = promedio de `Asistencia_Pct`.
- Aprobación: promedio ≥ 70 y asistencia ≥ 80.
- Participación: promedio < 70 y asistencia ≥ 80.
- Sin certificado: asistencia < 80.
