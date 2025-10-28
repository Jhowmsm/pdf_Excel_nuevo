from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import Dict, List, Tuple
import fitz  # PyMuPDF
import uuid
import json
import re
from openpyxl import Workbook

app = FastAPI(
    title="PDF Table Extractor (RAW line only)",
    description="Extrae cada fila de la tabla como una sola línea de texto en Excel.",
    version="2.1.0"
)

# ⚠ Ajusta dominios si tu codespace cambia de nombre
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

# Detectar NIF/NIE/CIF como "ancla fiscal"
PAT_ID_GENERAL = re.compile(
    r"""(
        [A-Z]\d{8}           |   # CIF tipo B26249896
        \d{8}[A-Z]           |   # NIF tipo 40313009N
        [XYZ]\d{7}[A-Z]          # NIE tipo X1234567A
    )""",
    re.VERBOSE
)

def parece_id_fiscal_en_linea(texto: str) -> bool:
    # devuelve True si encontramos algún identificador fiscal en la línea
    return bool(PAT_ID_GENERAL.search(texto))

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
    # Comparamos texto del PDF con el encabezado esperado (ignorando mayúsculas/tildes)
    norm_span = normalizar(span_text)
    norm_ref = normalizar(referencia)
    if norm_ref in norm_span or norm_span in norm_ref:
        return True
    palabras_ref = [p for p in norm_ref.split() if len(p) > 4]
    for p in palabras_ref:
        if p in norm_span:
            return True
    return False

@app.post("/procesar/")
async def procesar_pdf(
    file: UploadFile = File(...),
    referencias: str = Form(...),
    exclusiones: str = Form("{}"),
    pagina_inicio: int = Form(...),
    pagina_fin: int = Form(...)
):
    """
    Modo RAW simple:
    - Lee sólo páginas en el rango.
    - Empieza a capturar después de ver headers tipo:
      'NOMBRE Y APELLIDOS O RAZÓN SOCIAL', 'NIF', 'MARCA', 'MATRÍCULA'.
    - Para cada fila visual:
        - Ordena spans por X
        - Une TODO como una sola línea
        - Filtra ruido tipo '¿' o líneas vacías
    - Genera Excel con una sola columna: RAW_LINE
    """

    # 1. Cargar y parsear las referencias/exclusiones/envío del front
    try:
        referencias_list: List[str] = json.loads(referencias)
        exclusiones_map: Dict[str, List[str]] = json.loads(exclusiones)
    except Exception:
        return {"error": "No se pudo leer referencias o exclusiones"}

    # 2. Convertir páginas 1-based → 0-based
    start_page = max(pagina_inicio - 1, 0)
    end_page = max(pagina_fin - 1, start_page)

    # 3. Guardar PDF temporalmente
    original_name = file.filename or "archivo.pdf"
    temp_pdf = f"/tmp/{uuid.uuid4().hex}_{original_name}"
    with open(temp_pdf, "wb") as f:
        f.write(await file.read())

    # 4. Abrir con PyMuPDF
    doc = fitz.open(temp_pdf)
    end_page = min(end_page, len(doc) - 1)

    TOL_Y = 4  # tolerancia vertical para agrupar spans en una fila visual
    filas_raw: List[str] = []

    en_modo_tabla = False  # se activa cuando detectamos encabezados

    # 5. Recorremos el rango de páginas
    for pageno in range(start_page, end_page + 1):
        page = doc[pageno]
        page_dict = page.get_text("dict")

        # 5.1 Detectar headers en esta página
        pagina_tiene_header = False
        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    texto_span = span["text"].strip()
                    for ref in referencias_list:
                        if header_match(texto_span, ref):
                            pagina_tiene_header = True
                            break

        if pagina_tiene_header:
            en_modo_tabla = True

        if not en_modo_tabla:
            # si aún no apareció cabecera de tabla, saltamos esta página
            continue

        # 5.2 Agrupar spans por fila en esta página
        filas_pag_cruda: List[Tuple[float, List[dict]]] = []

        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                # centro vertical aproximado de esta línea
                ys = [s["bbox"][1] for s in spans] + [s["bbox"][3] for s in spans]
                y_centro = (min(ys) + max(ys)) / 2.0

                # buscar fila existente con Y cercana
                fila_idx = None
                for idx, (y_exist, _) in enumerate(filas_pag_cruda):
                    if abs(y_exist - y_centro) <= TOL_Y:
                        fila_idx = idx
                        break
                if fila_idx is None:
                    filas_pag_cruda.append((y_centro, []))
                    fila_idx = len(filas_pag_cruda) - 1

                for s in spans:
                    texto_span = s["text"].strip()
                    if not texto_span:
                        continue

                    # ignorar numeritos de paginación suelta (ej. "51")
                    if re.fullmatch(r"\d{1,3}", texto_span):
                        continue

                    # ignorar headers repetidos
                    if any(header_match(texto_span, ref) for ref in referencias_list):
                        continue

                    # ignorar exclusiones explícitas definidas por el usuario
                    saltar = False
                    for ref in referencias_list:
                        if texto_span in exclusiones_map.get(ref, []):
                            saltar = True
                            break
                    if saltar:
                        continue

                    x0, y0, x1, y1 = s["bbox"]
                    filas_pag_cruda[fila_idx][1].append({
                        "text": texto_span,
                        "x0": x0
                    })

        # 5.3 Transformar cada fila cruda en una línea plana
        for (_, spans_de_fila) in filas_pag_cruda:
            if not spans_de_fila:
                continue

            # ordenar por X izquierda→derecha
            spans_sorted = sorted(spans_de_fila, key=lambda sp: sp["x0"])
            tokens = [sp["text"] for sp in spans_sorted]

            # unir todo
            raw_line = " ".join(tokens)
            # limpiar espacios múltiples
            raw_line = re.sub(r"\s+", " ", raw_line).strip()

            if not raw_line:
                continue

            # Filtrar ruido:
            # - si es muy corto (1-2 chars) ignoramos
            if len(raw_line) < 4:
                continue

            # - si no tiene ningún ID fiscal Y además parece ser basura tipo "¿"
            #   (por ejemplo una sola palabra minúscula rara), también ignoramos
            if not parece_id_fiscal_en_linea(raw_line):
                # medir cuántas "palabras decentes" tiene
                palabras = [p for p in raw_line.split() if len(p) > 2]
                if len(palabras) < 2:
                    # ej "¿" o "i" o "¿ i"
                    continue

            # si pasó los filtros, la guardamos
            filas_raw.append(raw_line)

    doc.close()

    # 6. Construir Excel con una sola columna
    wb = Workbook()
    ws = wb.active
    ws.title = "FILAS_RAW"

    ws.append(["RAW_LINE"])

    for linea in filas_raw:
        ws.append([linea])

    # Nota: ya NO incluimos NIE_Warnings ni columnas separadas,
    # porque la idea ahora es que tú hagas el split lógico en Google Sheets.

    output_path = f"/tmp/resultado_{uuid.uuid4().hex}.xlsx"
    wb.save(output_path)

    return FileResponse(
        output_path,
        filename="resultado.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
