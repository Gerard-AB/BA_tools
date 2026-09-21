"""
Web para convertir el fichero de inscripciones en el Excel de pagos.

Arrancar con:  python -m uvicorn app:app --reload
Luego abrir:   http://127.0.0.1:8000
"""

import base64
from pathlib import Path

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from procesador import CABECERA, ErrorProceso, convertir

ESTATICOS = Path(__file__).parent / "static"
MAX_BYTES = 10 * 1024 * 1024  # 10 MB

app = FastAPI(title="Pagos Bons Amics")


@app.get("/")
def inicio():
    return FileResponse(ESTATICOS / "index.html")


@app.get("/logo.png")
def logo():
    return FileResponse(ESTATICOS / "logo.png", media_type="image/png")


@app.post("/procesar")
async def procesar_fichero(fichero: UploadFile = File(...)):
    contenido = await fichero.read()

    if not contenido:
        return JSONResponse({"error": "El fichero está vacío."}, status_code=400)
    if len(contenido) > MAX_BYTES:
        return JSONResponse({"error": "El fichero supera los 10 MB."}, status_code=400)

    try:
        xlsx, hojas, avisos = convertir(contenido, fichero.filename or "")
    except ErrorProceso as e:
        return JSONResponse({"error": str(e)}, status_code=400)
    except Exception as e:
        return JSONResponse({"error": f"Error inesperado: {e}"}, status_code=500)

    nombre_salida = Path(fichero.filename or "resultado").stem + "_pagos.xlsx"

    return {
        "nombre": nombre_salida,
        "cabecera": CABECERA,
        "hojas": hojas,
        "total_filas": sum(len(h["filas"]) for h in hojas),
        "total_importe": sum(h["total"] for h in hojas),
        "avisos": avisos,
        "xlsx": base64.b64encode(xlsx).decode("ascii"),
    }
