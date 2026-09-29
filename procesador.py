"""
Lógica de conversión: fichero de inscripciones (XLSX o CSV) -> un XLSX por grupo.

Trabaja siempre en memoria, para poder usarla desde la web sin tocar el disco.
"""

import csv
import io
import unicodedata
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CABECERA = ["Nombre", "Apellidos", "Curso", "Teléfono", "Dirección", "Inscripción", "Pago"]

# Nombres aceptados para cada columna de entrada (ya normalizados: sin acentos y en minúsculas).
ALIAS = {
    "beneficiari": ["beneficiari", "beneficiario", "alumne", "alumno"],
    "classe": ["classe", "clase", "curso", "curs"],
    "total": ["total compra", "total", "import", "importe"],
}

FORMATO_EURO = '#,##0.00\\ "€"'

# Grupos de curso -> hoja propia con su color de cabecera.
# Cada grupo se identifica por la etapa (primaria / eso) y el nivel.
GRUPOS = [
    {
        "clave": "primaria_34",
        "hoja": "Primària 3-4",
        "etapa": "primaria",
        "niveles": {"tercer", "tercero", "quart", "cuarto"},
        "color": "2E7D32",          # verde
        "color_texto": "FFFFFF",
    },
    {
        "clave": "primaria_56",
        "hoja": "Primària 5-6",
        "etapa": "primaria",
        "niveles": {"cinque", "cinquen", "quinto", "sise", "sisen", "sexto"},
        "color": "7B4EA3",          # lila
        "color_texto": "FFFFFF",
    },
    {
        "clave": "eso_12",
        "hoja": "ESO 1-2",
        "etapa": "eso",
        "niveles": {"primer", "primero", "segon", "segundo"},
        "color": "FFC000",          # amarillo
        "color_texto": "1B2430",
    },
]

GRUPO_RESTO = {
    "clave": "otros",
    "hoja": "Otros",
    "color": "1F4E79",              # azul
    "color_texto": "FFFFFF",
}


class ErrorProceso(Exception):
    """Error esperable que se muestra tal cual al usuario."""


def normalizar(texto):
    """
    Quita acentos, pasa a minúsculas y elimina espacios sobrantes.
    Se usa solo para comparar cabeceras, no para los datos.
    """
    if texto is None:
        return ""
    sin_acentos = ''.join(
        c for c in unicodedata.normalize('NFKD', str(texto))
        if not unicodedata.combining(c)
    )
    return sin_acentos.replace('﻿', '').strip().lower()


def _texto(valor):
    """Convierte un valor de celda a texto limpio."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def leer_filas(contenido, nombre_fichero):
    """
    Devuelve las filas del fichero subido como lista de listas de texto.
    Acepta .xlsx / .xlsm y .csv / .txt.
    """
    extension = nombre_fichero.lower().rsplit(".", 1)[-1] if "." in nombre_fichero else ""

    if extension in ("xlsx", "xlsm"):
        return _leer_xlsx(contenido)
    if extension in ("csv", "txt", "tsv"):
        return _leer_csv(contenido)
    if extension == "xls":
        raise ErrorProceso(
            "El formato .xls (Excel antiguo) no está soportado. "
            "Ábrelo en Excel y guárdalo como .xlsx."
        )
    raise ErrorProceso(f"Formato no soportado: .{extension}. Sube un .xlsx o un .csv.")


def _leer_xlsx(contenido):
    try:
        wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    except Exception as e:
        raise ErrorProceso(f"No se ha podido abrir el Excel: {e}")
    try:
        ws = wb.active
        return [[_texto(celda) for celda in fila] for fila in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _leer_csv(contenido):
    texto = None
    for codificacion in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            texto = contenido.decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    if texto is None:
        raise ErrorProceso("No se ha podido leer el CSV: codificación desconocida.")

    primera_linea = texto.split("\n", 1)[0]
    delimitador = ";" if primera_linea.count(";") >= primera_linea.count(",") else ","

    lector = csv.reader(io.StringIO(texto), delimiter=delimitador)
    return [[_texto(c) for c in fila] for fila in lector]


def _buscar_indices(cabeceras):
    """Localiza las columnas de entrada aceptando varios nombres posibles."""
    normalizadas = [normalizar(c) for c in cabeceras]
    indices = {}
    for clave, nombres in ALIAS.items():
        for nombre in nombres:
            if nombre in normalizadas:
                indices[clave] = normalizadas.index(nombre)
                break
        else:
            encontradas = ", ".join(c for c in cabeceras if c) or "(ninguna)"
            raise ErrorProceso(
                f"No se ha encontrado la columna «{nombres[0]}». "
                f"Cabeceras detectadas: {encontradas}"
            )
    return indices


def clasificar(curso):
    """
    Decide a qué grupo pertenece un curso a partir de su etapa y su nivel.
    «Primer de la ESO» y «Primer de Primària» comparten nivel, por eso se mira
    también la etapa. Lo que no encaje en ningún grupo va a «Otros».
    """
    palabras = set(
        p for p in ''.join(c if c.isalnum() else ' ' for c in normalizar(curso)).split() if p
    )
    etapa = "eso" if "eso" in palabras else ("primaria" if "primaria" in palabras else None)

    if etapa:
        for grupo in GRUPOS:
            if grupo["etapa"] == etapa and palabras & grupo["niveles"]:
                return grupo
    return GRUPO_RESTO


def _parsear_importe(texto):
    """
    Convierte '45', '54,50', '1.234,56 €' -> float. Devuelve None si no es un número.
    """
    limpio = texto.replace("€", "").replace(" ", "").replace("\xa0", "")
    if not limpio:
        return None
    if "," in limpio:  # formato español: el punto es separador de miles
        limpio = limpio.replace(".", "").replace(",", ".")
    try:
        return float(limpio)
    except ValueError:
        return None


def procesar(filas):
    """
    Aplica la transformación y devuelve (hojas, avisos).

    Cada hoja es un grupo de curso con su color y sus filas; una fila es
    [nombre, apellidos, curso, "", "", "", importe].
    """
    filas = [f for f in filas if any(c for c in f)]
    if not filas:
        raise ErrorProceso("El fichero está vacío.")

    indices = _buscar_indices(filas[0])
    minimo = max(indices.values()) + 1

    por_grupo = {}
    avisos = []
    for numero, fila in enumerate(filas[1:], start=2):
        if len(fila) < minimo:
            continue

        beneficiari = fila[indices["beneficiari"]].strip()
        if not beneficiari:
            continue

        if "," in beneficiari:
            apellidos, nombre = [parte.strip() for parte in beneficiari.split(",", 1)]
        else:
            nombre, apellidos = beneficiari, ""
            avisos.append(f"Fila {numero}: «{beneficiari}» no tiene coma; va entero en Nombre.")

        curso = fila[indices["classe"]].strip()
        bruto = fila[indices["total"]].strip()
        importe = _parsear_importe(bruto)
        if importe is None and bruto:
            avisos.append(f"Fila {numero}: importe «{bruto}» no es un número; se copia tal cual.")

        grupo = clasificar(curso)
        por_grupo.setdefault(grupo["clave"], []).append(
            [nombre, apellidos, curso, "", "", "", importe if importe is not None else bruto]
        )

    if not por_grupo:
        raise ErrorProceso("No se ha encontrado ninguna fila de datos válida.")

    # Orden fijo de hojas: los tres grupos conocidos y, al final, «Otros»
    hojas = []
    for grupo in GRUPOS + [GRUPO_RESTO]:
        filas_grupo = por_grupo.get(grupo["clave"])
        if not filas_grupo:
            continue
        hojas.append({
            "clave": grupo["clave"],
            "nombre": grupo["hoja"],
            "color": grupo["color"],
            "color_texto": grupo["color_texto"],
            "filas": filas_grupo,
            "total": sum(f[6] for f in filas_grupo if isinstance(f[6], (int, float))),
        })

    return hojas, avisos


def _montar_hoja(ws, hoja):
    """Vuelca una hoja con su cabecera de color, formato de moneda y anchos."""
    ws.append(CABECERA)
    for fila in hoja["filas"]:
        ws.append(fila)

    # Cabecera con el color del grupo, y fijada al hacer scroll
    relleno = PatternFill("solid", fgColor=hoja["color"])
    for celda in ws[1]:
        celda.font = Font(bold=True, color=hoja["color_texto"])
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    # Formato de moneda en la columna Pago
    columna_pago = len(CABECERA)
    for fila in ws.iter_rows(min_row=2, min_col=columna_pago, max_col=columna_pago):
        for celda in fila:
            if isinstance(celda.value, (int, float)):
                celda.number_format = FORMATO_EURO

    # Ancho de columnas según el contenido
    for i in range(1, len(CABECERA) + 1):
        ancho = max(
            (len(_texto(c.value)) for c in ws[get_column_letter(i)] if c.value is not None),
            default=10,
        )
        ws.column_dimensions[get_column_letter(i)].width = min(max(ancho + 2, 10), 40)


def generar_xlsx(hoja):
    """Construye el XLSX de un grupo en memoria y devuelve sus bytes."""
    wb = Workbook()
    ws = wb.active
    ws.title = hoja["nombre"]
    ws.sheet_properties.tabColor = hoja["color"]
    _montar_hoja(ws, hoja)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def nombre_archivo(base, hoja):
    """
    Nombre del fichero de un grupo: «listado_primaria-3-4.xlsx».
    Sin acentos ni espacios, para que no se rompa al descargar.
    """
    trozos = ''.join(
        c if c.isalnum() else ' ' for c in normalizar(hoja["nombre"])
    ).split()
    return f"{base}_{'-'.join(trozos)}.xlsx"


def convertir(contenido, nombre_fichero):
    """
    Punto de entrada: bytes subidos -> (hojas, avisos).

    Cada hoja lleva ya su propio XLSX en la clave «xlsx», de modo que se
    descarga un fichero por grupo en vez de uno solo con varias pestañas.
    """
    base = Path(nombre_fichero or "resultado").stem or "resultado"
    filas_origen = leer_filas(contenido, nombre_fichero)
    hojas, avisos = procesar(filas_origen)

    for hoja in hojas:
        hoja["archivo"] = nombre_archivo(base, hoja)
        hoja["xlsx"] = generar_xlsx(hoja)

    return hojas, avisos
