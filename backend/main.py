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


app = FastAPI(
    title="PDF Table Extractor",
    description="Extrae columnas de tablas en PDF y las exporta a Excel.",
    version="1.0.0"
)

# ⬇⬇⬇ AJUSTA ESTO si cambias de codespace ⬇⬇⬇
# Usa exactamente tus dos URLs públicas (backend:8001, frontend:5500).
# Si tu codespace cambia de nombre, reemplaza weary-cape-v6gp9gxg7jj52wqpw-... por el nuevo prefijo.
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://weary-cape-v6gp9gxg7jj52wqpw-8001.app.github.dev",  # backend público
        "https://weary-cape-v6gp9gxg7jj52wqpw-5500.app.github.dev",  # frontend público

        # extras para desarrollo local o futuro despliegue
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


def normalizar(s: str) -> str:
    """
    Normaliza texto (acentos fuera, minúsculas, limpia signos y espacios).
    Esto nos sirve para reconocer los encabezados de las columnas aunque el PDF
    traiga variaciones de mayúsculas y tildes.
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
    Devuelve True si el texto del PDF parece ser el encabezado 'referencia'.
    Permitimos diferencias de formato (tildes, mayúsculas, barras, etc.).
    """
    norm_span = normalizar(span_text)
    norm_ref = normalizar(referencia)

    # match directo o contenido
    if norm_ref in norm_span or norm_span in norm_ref:
        return True

    # Fallback: palabras "largas" (>4 letras)
    palabras_ref = [p for p in norm_ref.split() if len(p) > 4]
    for p in palabras_ref:
        if p in norm_span:
            return True

    return False


# regex para detectar NIF / NIE en todo el documento
PAT_ID = re.compile(
    r"""(
        [XYZ]\d{7}[A-Z]      # NIE tipo X/Y/Z + 7 dígitos + letra
        |
        \d{8}[A-Z]           # NIF clásico: 8 dígitos + letra
    )""",
    re.VERBOSE
)


@app.post("/procesar/")
async def procesar_pdf(
    file: UploadFile = File(...),
    referencias: str = Form(...),
    exclusiones: str = Form("{}")
):
    """
    Espera:
    - file: PDF subido
    - referencias: lista JSON de encabezados de columnas
    - exclusiones: diccionario JSON con textos a ignorar por columna

    Devuelve:
    - Un Excel con la extracción
    """

    # 1. Parsear los JSON recibidos
    try:
        referencias_list: List[str] = json.loads(referencias)
        exclusiones_map: Dict[str, List[str]] = json.loads(exclusiones)
    except Exception:
        return {"error": "No se pudo leer referencias o exclusiones"}

    # 2. Guardar el PDF temporalmente
    original_name = file.filename or "archivo.pdf"
    temp_pdf = f"/tmp/{uuid.uuid4().hex}_{original_name}"
    with open(temp_pdf, "wb") as f:
        f.write(await file.read())

    # 3. Abrir PDF
    doc = fitz.open(temp_pdf)

    # 4. Buscar posibles NIF/NIE globales para la hoja "NIE_Warnings"
    ids_detectados = set()
    for page in doc:
        page_text = page.get_text("text")
        for m in PAT_ID.findall(page_text):
            ids_detectados.add(m)

    # 5. Recorrer las páginas, identificar columnas y agrupar filas
    MARGEN_X = 25       # tolerancia horizontal al asignar un span a una columna
    TOL_Y = 4           # tolerancia vertical para agrupar spans en la misma fila
    todas_filas: List[Dict[str, str]] = []

    # Coordenadas X de las columnas conocidas (se va llenando y recordando entre páginas)
    ref_x_global: Dict[str, float] = {}

    for page in doc:
        page_dict = page.get_text("dict")

        # 5.1. Detectar encabezados de esta página y guardar su coordenada X inicial
        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    texto_span = span["text"].strip()
                    x0, y0, x1, y1 = span["bbox"]

                    for ref in referencias_list:
                        # si aún no tenemos X fija para esa referencia
                        if ref not in ref_x_global and header_match(texto_span, ref):
                            ref_x_global[ref] = x0

        # 5.2. Agrupar spans de texto en filas visuales (por coordenada Y)
        # cada elemento: (y_centro, {columnaRef: textoAcumulado})
        filas_pag: List[Tuple[float, Dict[str, str]]] = []

        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                ys = [s["bbox"][1] for s in spans] + [s["bbox"][3] for s in spans]
                y_centro = (min(ys) + max(ys)) / 2.0

                # buscar si ya tenemos una fila "cercana" en Y
                fila_idx = None
                for idx, (y_exist, data_dict) in enumerate(filas_pag):
                    if abs(y_exist - y_centro) <= TOL_Y:
                        fila_idx = idx
                        break

                # si no existe, creamos nueva fila con todas las refs vacías
                if fila_idx is None:
                    filas_pag.append(
                        (y_centro, {ref: "" for ref in referencias_list})
                    )
                    fila_idx = len(filas_pag) - 1

                # Para cada span de esta línea: asignarlo a la columna más cercana en X
                for s in spans:
                    texto_span = s["text"].strip()
                    if not texto_span:
                        continue

                    # 1) saltar si este span ES un encabezado
                    es_header = any(header_match(texto_span, ref) for ref in referencias_list)
                    if es_header:
                        continue

                    # 2) saltar si está en exclusiones
                    en_exclusion = False
                    for ref in referencias_list:
                        if texto_span in exclusiones_map.get(ref, []):
                            en_exclusion = True
                            break
                    if en_exclusion:
                        continue

                    # 3) saltar si es un número de página suelto ("51")
                    if re.fullmatch(r"\d{1,3}", texto_span):
                        continue

                    x0, y0, x1, y1 = s["bbox"]

                    # decidir columna por distancia horizontal
                    mejor_ref = None
                    mejor_dist = None

                    for ref, ref_x in ref_x_global.items():
                        dist = abs(x0 - ref_x)
                        if dist <= MARGEN_X and (mejor_dist is None or dist < mejor_dist):
                            mejor_dist = dist
                            mejor_ref = ref

                    if mejor_ref is not None:
                        actual = filas_pag[fila_idx][1][mejor_ref]
                        if actual:
                            nuevo = actual + " " + texto_span
                        else:
                            nuevo = texto_span
                        filas_pag[fila_idx][1][mejor_ref] = nuevo

        # 5.3. Agregar filas no vacías de esta página al total
        for y_centro, data_dict in filas_pag:
            if any(v.strip() for v in data_dict.values()):
                todas_filas.append(data_dict)

    doc.close()

    # 6. Crear Excel de salida
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    # encabezados
    ws.append(referencias_list)

    # estilo para marcar celdas largas en "Nombre / Razón Social"
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for fila_dict in todas_filas:
        row_vals = [fila_dict.get(ref, "") for ref in referencias_list]
        ws.append(row_vals)

        current_row_idx = ws.max_row
        for col_idx, ref in enumerate(referencias_list, start=1):
            # marcamos en amarillo si el nombre/razón social parece "demasiado largo"
            ref_norm = ref.lower()
            if "razón social" in ref_norm or "razon social" in ref_norm or ref_norm.startswith("nombre"):
                val = ws.cell(current_row_idx, col_idx).value or ""
                if len(val) > 60:
                    ws.cell(current_row_idx, col_idx).fill = amarillo

    # 7. Hoja de advertencias
    if ids_detectados:
        ws_alerta = wb.create_sheet("NIE_Warnings")
        ws_alerta.append(["NIF/NIE Detectados"])
        for codigo in sorted(ids_detectados):
            ws_alerta.append([codigo])

    # 8. Guardar archivo temporal y enviarlo
    output_path = f"/tmp/resultado_{uuid.uuid4().hex}.xlsx"
    wb.save(output_path)

    return FileResponse(
        output_path,
        filename="resultado.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
