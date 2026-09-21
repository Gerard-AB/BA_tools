"""
Tests de la lógica de conversión.

Ejecutar con:  python -m pytest
"""

import base64
import io
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import app  # noqa: E402
from procesador import CABECERA, ErrorProceso, clasificar, convertir  # noqa: E402

EJEMPLO = Path(__file__).resolve().parents[1] / "ejemplo" / "inscripciones-ejemplo.csv"

CSV_MINIMO = (
    "Beneficiari;Classe;Total compra\n"
    "Gómez Puig, Ada;Quart de Primària-A;45\n"
)


def csv_bytes(texto):
    return texto.encode("utf-8")


# --- Clasificación de cursos -------------------------------------------------

@pytest.mark.parametrize("curso, hoja", [
    ("Tercer de Primària-A", "Primària 3-4"),
    ("Quart de Primària-C", "Primària 3-4"),
    ("Cinqué de Primària-B", "Primària 5-6"),
    ("Sisè de Primària", "Primària 5-6"),
    ("Primer de la ESO-A", "ESO 1-2"),
    ("Segón de la ESO", "ESO 1-2"),
    ("Segon d'ESO-B", "ESO 1-2"),
    # El ordinal solo no basta: hay que mirar también la etapa.
    ("Primer de Primària", "Otros"),
    ("Tercer de la ESO", "Otros"),
    ("Infantil 5 anys", "Otros"),
    ("", "Otros"),
])
def test_clasificar(curso, hoja):
    assert clasificar(curso)["hoja"] == hoja


# --- Conversión --------------------------------------------------------------

def test_separa_nombre_y_apellidos():
    _, hojas, _ = convertir(csv_bytes(CSV_MINIMO), "x.csv")
    assert hojas[0]["filas"][0][:3] == ["Ada", "Gómez Puig", "Quart de Primària-A"]


def test_sin_coma_todo_va_al_nombre_y_avisa():
    _, hojas, avisos = convertir(
        csv_bytes("Beneficiari;Classe;Total compra\nSinComa;Quart de Primària;45\n"), "x.csv")
    assert hojas[0]["filas"][0][:2] == ["SinComa", ""]
    assert len(avisos) == 1


@pytest.mark.parametrize("bruto, esperado", [
    ("45", 45), ("54,50", 54.5), ("1.234,56 €", 1234.56), ("66,5", 66.5),
])
def test_importes(bruto, esperado):
    _, hojas, _ = convertir(
        csv_bytes(f"Beneficiari;Classe;Total compra\nApe, Nom;Quart de Primària;{bruto}\n"), "x.csv")
    assert hojas[0]["filas"][0][6] == esperado


def test_importe_no_numerico_se_copia_y_avisa():
    _, hojas, avisos = convertir(
        csv_bytes("Beneficiari;Classe;Total compra\nApe, Nom;Quart de Primària;gratis\n"), "x.csv")
    assert hojas[0]["filas"][0][6] == "gratis"
    assert len(avisos) == 1


def test_cabeceras_con_acentos_mayusculas_y_alias():
    _, hojas, _ = convertir(
        csv_bytes("BENEFICIARIO;Curso;Importe\nApe, Nom;Quart de Primària;45\n"), "x.csv")
    assert hojas[0]["filas"][0][0] == "Nom"


def test_csv_y_xlsx_dan_el_mismo_resultado():
    crudo = EJEMPLO.read_bytes()
    _, desde_csv, _ = convertir(crudo, "ejemplo.csv")

    wb = Workbook()
    ws = wb.active
    for linea in crudo.decode("utf-8-sig").splitlines():
        ws.append(linea.split(";"))
    buffer = io.BytesIO()
    wb.save(buffer)
    _, desde_xlsx, _ = convertir(buffer.getvalue(), "ejemplo.xlsx")

    assert [h["filas"] for h in desde_csv] == [h["filas"] for h in desde_xlsx]


# --- Errores esperados -------------------------------------------------------

@pytest.mark.parametrize("contenido, nombre, trozo", [
    (b"", "x.csv", "vac"),
    (b"a;b;c\n1;2;3\n", "x.csv", "No se ha encontrado la columna"),
    (b"lo que sea", "x.pdf", "Formato no soportado"),
    (b"lo que sea", "x.xls", "Excel antiguo"),
])
def test_errores(contenido, nombre, trozo):
    with pytest.raises(ErrorProceso, match=trozo):
        convertir(contenido, nombre)


# --- XLSX generado -----------------------------------------------------------

def test_xlsx_una_hoja_por_grupo_con_su_color():
    xlsx, hojas, _ = convertir(EJEMPLO.read_bytes(), "ejemplo.csv")
    wb = load_workbook(io.BytesIO(xlsx))

    assert wb.sheetnames == ["Primària 3-4", "Primària 5-6", "ESO 1-2", "Otros"]

    for hoja in hojas:
        ws = wb[hoja["nombre"]]
        assert [c.value for c in ws[1]] == CABECERA
        assert ws.max_row == len(hoja["filas"]) + 1
        assert ws.cell(1, 1).fill.fgColor.rgb[-6:] == hoja["color"]
        assert ws.freeze_panes == "A2"
        assert "€" in ws.cell(2, len(CABECERA)).number_format


# --- API ---------------------------------------------------------------------

def test_endpoint_devuelve_el_excel():
    cliente = TestClient(app)
    respuesta = cliente.post(
        "/procesar", files={"fichero": ("ejemplo.csv", EJEMPLO.read_bytes(), "text/csv")})

    assert respuesta.status_code == 200
    datos = respuesta.json()
    assert datos["nombre"] == "ejemplo_pagos.xlsx"
    assert datos["total_filas"] == 20
    assert len(datos["hojas"]) == 4
    assert load_workbook(io.BytesIO(base64.b64decode(datos["xlsx"]))).sheetnames


def test_endpoint_rechaza_fichero_invalido():
    cliente = TestClient(app)
    respuesta = cliente.post(
        "/procesar", files={"fichero": ("x.pdf", b"algo", "application/pdf")})

    assert respuesta.status_code == 400
    assert "error" in respuesta.json()


def test_pagina_y_logo_se_sirven():
    cliente = TestClient(app)
    assert cliente.get("/").status_code == 200
    assert cliente.get("/logo.png").headers["content-type"] == "image/png"
