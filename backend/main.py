from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import Dict, List, Tuple
import fitz  # PyMuPDF
import uuid
import json
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill

app = FastAPI(
    title="PDF Table Extractor",
    description="Extrae columnas de tablas en PDF y las exporta a Excel (rango de páginas).",
    version="1.2.0"
)

# ⚠ Ajusta estos dominios si el codespace cambia de nombre
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

def normalizar(s: str) -> str:
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
    norm_span = normalizar(span_text)
    norm_ref = normalizar(referencia)
    if norm_ref in norm_span or norm_span in norm_ref:
        return True
    palabras_ref = [p for p in norm_ref.split() if len(p) > 4]
    for p in palabras_ref:
        if p in norm_span:
            return True
    return False

PAT_ID = re.compile(
    r"""(
        [XYZ]\d{7}[A-Z]      # NIE X/Y/Z
        |
        \d{8}[A-Z]           # NIF 12345678Z
    )""",
    re.VERBOSE
)

@app.post("/procesar/")
async def procesar_pdf(
    file: UploadFile = File(...),
    referencias: str = Form(...),
    exclusiones: str = Form("{}"),
    pagina_inicio: int = Form(...),
    pagina_fin: int = Form(...)
):
    """
    pagina_inicio y pagina_fin vienen 1-based desde el frontend.
    Vamos a convertirlos a índices 0-based internamente.
    Solo procesamos ese rango.
    """

    try:
        referencias_list: List[str] = json.loads(referencias)
        exclusiones_map: Dict[str, List[str]] = json.loads(exclusiones)
    except Exception:
        return {"error": "No se pudo leer referencias o exclusiones"}

    # Ajustar páginas a base 0
    start_page = max(pagina_inicio - 1, 0)
    end_page = max(pagina_fin - 1, start_page)

    original_name = file.filename or "archivo.pdf"
    temp_pdf = f"/tmp/{uuid.uuid4().hex}_{original_name}"
    with open(temp_pdf, "wb") as f:
        f.write(await file.read())

    doc = fitz.open(temp_pdf)

    # limitar a rango válido
    end_page = min(end_page, len(doc) - 1)

    MARGEN_X = 25
    TOL_Y = 4

    todas_filas: List[Dict[str, str]] = []
    ref_x_global: Dict[str, float] = {}

    # Detectar NIF/NIE solo en el rango de páginas
    ids_detectados = set()
    for pageno in range(start_page, end_page + 1):
        page = doc[pageno]
        page_text = page.get_text("text")
        for m in PAT_ID.findall(page_text):
            ids_detectados.add(m)

    # Bandera igual que antes: no empiezo a capturar hasta ver headers
    en_modo_tabla = False

    for pageno in range(start_page, end_page + 1):
        page = doc[pageno]
        page_dict = page.get_text("dict")

        pagina_tiene_header = False

        # 1. Detectar headers en ESTA página dentro del rango
        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    texto_span = span["text"].strip()
                    x0, y0, x1, y1 = span["bbox"]

                    for ref in referencias_list:
                        if header_match(texto_span, ref):
                            pagina_tiene_header = True
                            if ref not in ref_x_global:
                                ref_x_global[ref] = x0

        if pagina_tiene_header:
            en_modo_tabla = True

        if not en_modo_tabla:
            continue

        # 2. Agrupar spans en filas visuales
        filas_pag: List[Tuple[float, Dict[str, str]]] = []

        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                ys = [s["bbox"][1] for s in spans] + [s["bbox"][3] for s in spans]
                y_centro = (min(ys) + max(ys)) / 2.0

                fila_idx = None
                for idx, (y_exist, data_dict) in enumerate(filas_pag):
                    if abs(y_exist - y_centro) <= TOL_Y:
                        fila_idx = idx
                        break

                if fila_idx is None:
                    filas_pag.append(
                        (y_centro, {ref: "" for ref in referencias_list})
                    )
                    fila_idx = len(filas_pag) - 1

                for s in spans:
                    texto_span = s["text"].strip()
                    if not texto_span:
                        continue

                    # saltar header repetido
                    if any(header_match(texto_span, ref) for ref in referencias_list):
                        continue

                    # saltar exclusiones
                    exclu = False
                    for ref in referencias_list:
                        if texto_span in exclusiones_map.get(ref, []):
                            exclu = True
                            break
                    if exclu:
                        continue

                    # saltar numeritos tipo nº de página
                    if re.fullmatch(r"\d{1,3}", texto_span):
                        continue

                    x0, y0, x1, y1 = s["bbox"]

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

        # guardar filas de esta página
        for y_centro, data_dict in filas_pag:
            if any(v.strip() for v in data_dict.values()):
                todas_filas.append(data_dict)

    doc.close()

    # 3. Generar Excel con los resultados
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    ws.append(referencias_list)

    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for fila_dict in todas_filas:
        row_vals = [fila_dict.get(ref, "") for ref in referencias_list]
        ws.append(row_vals)

        current_row_idx = ws.max_row
        for col_idx, ref in enumerate(referencias_list, start=1):
            ref_norm = ref.lower()
            if "razón social" in ref_norm or "razon social" in ref_norm or ref_norm.startswith("nombre"):
                val = ws.cell(current_row_idx, col_idx).value or ""
                if len(val) > 60:
                    ws.cell(current_row_idx, col_idx).fill = amarillo

    if ids_detectados:
        ws_alerta = wb.create_sheet("NIE_Warnings")
        ws_alerta.append(["NIF/NIE Detectados"])
        for codigo in sorted(ids_detectados):
            ws_alerta.append([codigo])

    output_path = f"/tmp/resultado_{uuid.uuid4().hex}.xlsx"
    wb.save(output_path)

    return FileResponse(
        output_path,
        filename="resultado.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
