from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import Dict, List, Tuple
import fitz  # PyMuPDF
import os
import uuid
import json
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# ------------------ FastAPI app ------------------

app = FastAPI(
    title="PDF Table Extractor",
    description="Extrae columnas de tablas en PDF y las exporta a Excel.",
    version="1.0.0"
)

# Habilita peticiones desde GitHub Pages y también localhost (para pruebas)

from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://weary-cape-v6gp9gxg7jj52wqpw-8001.app.github.dev",  # backend
        "https://weary-cape-v6gp9gxg7jj52wqpw-5500.app.github.dev",  # frontend
        "http://localhost:5500",
        "http://127.0.0.1:5500",
        "http://localhost:8000",
        "http://127.0.0.1:8000",
        "https://jhowmsm.github.io",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ------------------ Utilidades internas ------------------

def normalizar(s: str) -> str:
    """
    Normaliza texto para hacer matching flexible con encabezados:
    - pasa a minúsculas
    - quita tildes
    - reemplaza símbolos comunes por espacio
    - colapsa espacios múltiples
    """
    import unicodedata
    s = s.lower()
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    s = s.replace(":", " ").replace("/", " ").replace(".", " ").replace(",", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def header_match(span_text: str, referencia: str) -> bool:
    """
    Devuelve True si span_text parece ser el encabezado 'referencia'.
    Tolera mayúsculas, acentos, barras y espacios diferentes.
    """
    norm_span = normalizar(span_text)
    norm_ref = normalizar(referencia)

    # caso simple
    if norm_ref in norm_span or norm_span in norm_ref:
        return True

    # intenta por palabras 'fuertes' (>4 chars)
    palabras_ref = [p for p in norm_ref.split() if len(p) > 4]
    for p in palabras_ref:
        if p in norm_span:
            return True

    return False


# regex para capturar posibles NIF/NIE
PAT_ID = re.compile(
    r"""(
        [XYZ]\d{7}[A-Z]      # NIE tipo X/Y/Z + 7 dígitos + letra
        |
        \d{8}[A-Z]           # NIF clásico
    )""",
    re.VERBOSE
)

# ------------------ Endpoint principal ------------------

@app.post("/procesar/")
async def procesar_pdf(
    file: UploadFile = File(...),
    referencias: str = Form(...),
    exclusiones: str = Form("{}")
):
    """
    Recibe:
    - PDF
    - referencias: lista JSON con nombres de columnas a extraer
    - exclusiones: dict JSON { "REF": ["texto a ignorar", ...], ... }

    Devuelve:
    - Excel con datos + hoja de advertencias de identificadores.
    """

    # 1. Parsear inputs
    try:
        referencias_list: List[str] = json.loads(referencias)
        exclusiones_map: Dict[str, List[str]] = json.loads(exclusiones)
    except Exception:
        return {"error": "No se pudo leer referencias o exclusiones"}

    # 2. Guardar PDF temporalmente
    original_name = file.filename
    temp_pdf = f"/tmp/{uuid.uuid4().hex}_{original_name}"

    with open(temp_pdf, "wb") as f:
        f.write(await file.read())

    # 3. Abrir PDF
    doc = fitz.open(temp_pdf)

    # 4. Detectar NIF/NIE globalmente para la hoja de alerta
    ids_detectados = set()
    for page in doc:
        page_text = page.get_text("text")
        for m in PAT_ID.findall(page_text):
            ids_detectados.add(m)

    # 5. Recorrer páginas y construir filas
    MARGEN_X = 25        # tolerancia horizontal en px
    TOL_Y = 4            # tolerancia vertical para agrupar spans en misma fila
    todas_filas: List[Dict[str, str]] = []

    # ref_x_global recordará (entre páginas) la coordenada X de cada columna detectada
    ref_x_global: Dict[str, float] = {}

    for page in doc:
        page_dict = page.get_text("dict")

        # 5.1 detectar posibles cabeceras en ESTA página
        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    texto_span = span["text"].strip()
                    x0, y0, x1, y1 = span["bbox"]

                    for ref in referencias_list:
                        # solo fijar si aún no tenemos X para esa referencia
                        if ref not in ref_x_global and header_match(texto_span, ref):
                            ref_x_global[ref] = x0

        # 5.2 agrupar spans en filas visuales por coordenada Y
        # filas_pag = lista de tuplas: (y_centro, {ref: texto_acumulado})
        filas_pag: List[Tuple[float, Dict[str, str]]] = []

        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                # coordenada vertical promedio de la línea completa
                ys = [s["bbox"][1] for s in spans] + [s["bbox"][3] for s in spans]
                y_centro = (min(ys) + max(ys)) / 2.0

                # buscar si ya tenemos una fila cercana en Y
                fila_idx = None
                for idx, (y_exist, data_dict) in enumerate(filas_pag):
                    if abs(y_exist - y_centro) <= TOL_Y:
                        fila_idx = idx
                        break

                # si no existe, creamos una nueva fila con todas las refs inicializadas vacías
                if fila_idx is None:
                    filas_pag.append(
                        (y_centro, {ref: "" for ref in referencias_list})
                    )
                    fila_idx = len(filas_pag) - 1

                # Ahora asignar cada span a la mejor columna por X
                for s in spans:
                    texto_span = s["text"].strip()
                    if not texto_span:
                        continue

                    #  descartar cosas que claramente NO son datos
                    #  1) si el span parece ser el header
                    skip = False
                    for ref in referencias_list:
                        if header_match(texto_span, ref):
                            skip = True
                            break
                    if skip:
                        continue

                    #  2) descartar si está en la lista de exclusiones de alguna referencia
                    for ref in referencias_list:
                        if texto_span in exclusiones_map.get(ref, []):
                            skip = True
                            break
                    if skip:
                        continue

                    #  3) descartar números de página (p.ej. "51")
                    if re.fullmatch(r"\d{1,3}", texto_span):
                        continue

                    x0, y0, x1, y1 = s["bbox"]

                    # decidir a qué referencia pertenece este span por cercanía X
                    mejor_ref = None
                    mejor_dist = None
                    for ref, ref_x in ref_x_global.items():
                        dist = abs(x0 - ref_x)
                        if dist <= MARGEN_X and (mejor_dist is None or dist < mejor_dist):
                            mejor_dist = dist
                            mejor_ref = ref

                    if mejor_ref is not None:
                        # concatenar si ya hay texto en esa celda
                        actual = filas_pag[fila_idx][1][mejor_ref]
                        if actual:
                            nuevo = actual + " " + texto_span
                        else:
                            nuevo = texto_span
                        filas_pag[fila_idx][1][mejor_ref] = nuevo

        # 5.3 guardar filas no vacías
        for y_centro, data_dict in filas_pag:
            if any(v.strip() for v in data_dict.values()):
                todas_filas.append(data_dict)

    # cerramos PDF
    doc.close()

    # 6. Generar Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    # encabezados
    ws.append(referencias_list)

    # estilo para marcar filas potencialmente sospechosas
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for fila_dict in todas_filas:
        row_vals = [fila_dict.get(ref, "") for ref in referencias_list]
        ws.append(row_vals)

        current_row_idx = ws.max_row
        for col_idx, ref in enumerate(referencias_list, start=1):
            # ejemplo de marca: si la celda de "NOMBRE Y APELLIDOS O RAZÓN SOCIAL"
            # es muy larga (posible concatenación de más de una fila)
            if ref.lower().startswith("nombre") or "razón social" in ref.lower() or "razon social" in ref.lower():
                val = ws.cell(current_row_idx, col_idx).value or ""
                if len(val) > 60:  # umbral ajustable
                    ws.cell(current_row_idx, col_idx).fill = amarillo

    # Hoja de advertencias
    if ids_detectados:
        ws_alerta = wb.create_sheet("NIE_Warnings")
        ws_alerta.append(["NIF/NIE Detectados"])
        for codigo in sorted(ids_detectados):
            ws_alerta.append([codigo])

    # 7. Guardar archivo temporal y responderlo
    output_path = f"/tmp/resultado_{uuid.uuid4().hex}.xlsx"
    wb.save(output_path)

    return FileResponse(
        output_path,
        filename="resultado.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
