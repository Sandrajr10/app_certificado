# app_certificado — Academia Horizonte

Aplicación web en Flask que lee dos Excel, calcula el promedio y la asistencia de
cada estudiante y decide qué certificado le corresponde (Aprobación, Participación
o Sin certificado).

## Estructura

```
app.py                      # Toda la lógica (la "cocina")
requirements.txt            # Flask + openpyxl
iniciar_app.bat             # Arranque local con doble clic
templates/index.html        # Página (el "salón": solo muestra)
public/static/style.css     # Estilos (Vercel los sirve desde public/)
demo/                       # Excel ficticios para la demo pública
Insumos/                    # Excel reales (NO se suben: datos personales)
vercel.json                 # Configuración de despliegue
.vercelignore               # Evita subir Insumos/ y archivos innecesarios
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

## Reglas de negocio

- Se agrupa por `Identificacion + Programa`: un certificado por estudiante y programa.
- Promedio = suma de las notas / cantidad de módulos cursados.
- Asistencia = promedio de `Asistencia_Pct`.
- Aprobación: promedio ≥ 70 y asistencia ≥ 80.
- Participación: promedio < 70 y asistencia ≥ 80.
- Sin certificado: asistencia < 80.
