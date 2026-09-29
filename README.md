# Pagos Bons Amics

Aplicación web que convierte el listado de inscripciones del colegio en Excel de
pagos: separa nombre y apellidos, normaliza los importes y reparte a los alumnos por
ciclo, generando **un fichero por grupo** con su propio color.

Subes un `.xlsx` o un `.csv`, ves el resultado en pantalla y descargas cada grupo por
separado. El servidor no guarda nada en disco: todo se procesa en memoria y los
ficheros viajan de vuelta en la misma respuesta.

## Instalación

Requiere Python 3.10 o superior.

```bash
python -m venv .venv
.venv\Scripts\activate        # en Linux o macOS: source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
python -m uvicorn app:app --reload
```

Abre <http://127.0.0.1:8000> y arrastra el fichero. En `ejemplo/` hay un CSV con
datos inventados para probarlo.

## Qué hace la conversión

Del fichero de entrada solo se usan tres columnas: **Beneficiari**, **Classe** y
**Total compra**. La búsqueda ignora acentos y mayúsculas, y acepta variantes
(`Classe` / `Clase` / `Curso`, `Total compra` / `Total` / `Import`…).

| Columna de salida | De dónde sale |
| --- | --- |
| Nombre, Apellidos | `Beneficiari`, partido por la primera coma (`Apellidos, Nombre`) |
| Curso | `Classe`, tal cual |
| Teléfono, Dirección, Inscripción | Vacías, para rellenar a mano |
| Pago | `Total compra` como número, con formato de moneda |

El importe se escribe como número real, no como texto, para que las sumas de Excel
funcionen. Acepta coma decimal y separador de miles (`1.234,56 €` → `1234.56`).

### Un Excel por grupo

Cada alumno va al fichero de su ciclo. Solo se generan los grupos que tienen alumnos,
así que un listado de un solo ciclo produce un único Excel.

| Curso | Fichero | Color de cabecera |
| --- | --- | --- |
| Tercer o Quart de Primària | `<listado>_primaria-3-4.xlsx` | Verde `#2E7D32` |
| Cinqué o Sisé de Primària | `<listado>_primaria-5-6.xlsx` | Lila `#7B4EA3` |
| Primer o Segón de la ESO | `<listado>_eso-1-2.xlsx` | Amarillo `#FFC000` |
| Cualquier otro | `<listado>_otros.xlsx` | Azul `#1F4E79` |

El nombre de cada fichero parte del que subiste, así que `inscripciones-ejemplo.csv`
produce `inscripciones-ejemplo_primaria-3-4.xlsx` y compañía.

La clasificación mira **etapa y nivel por separado**, porque `Primer de la ESO` y
`Primer de Primària` comparten ordinal y van a hojas distintas.

Si alguna fila no se puede interpretar del todo —un beneficiario sin coma, un importe
que no es un número— la fila se procesa igual y la web avisa indicando cuál, para que
la revises.

## Estructura

```
app.py          API FastAPI: sirve la página y expone POST /procesar
procesador.py   Toda la lógica de conversión, sin dependencias de web
static/         index.html (página) y logo.png
ejemplo/        CSV de ejemplo con datos inventados
tests/          Tests de la conversión y de la API
```

`procesador.py` no importa nada de FastAPI, así que la conversión se puede usar desde
un script o desde otro programa sin levantar el servidor:

```python
from procesador import convertir

hojas, avisos = convertir(open("listado.csv", "rb").read(), "listado.csv")
for hoja in hojas:
    open(hoja["archivo"], "wb").write(hoja["xlsx"])
```

## Tests

```bash
pip install -r requirements-dev.txt
python -m pytest
```

## Datos personales

Los listados de inscripciones contienen **nombres de alumnos menores de edad**. No los
subas al repositorio: `.gitignore` excluye todos los `.csv` y `.xlsx` salvo los de
`ejemplo/`, y la carpeta `datos/` está ignorada entera. Úsala para tus ficheros reales.

Ten en cuenta que en Git un fichero subido por error sigue en el historial aunque luego
lo borres en un commit posterior.
