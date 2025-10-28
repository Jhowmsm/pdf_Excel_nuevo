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
    description="Extrae columnas PDF → Excel usando detección de NIF como ancla.",
    version="1.4.0"
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

# Regex general para IDs fiscales españoles típicos:
# - CIF: letra + 8 dígitos  (B26249896)
# - NIF: 8 dígitos + letra  (40313009N)
# - NIE: [XYZ] + 7 dígitos + letra (X1234567A)
PAT_ID_GENERAL = re.compile(
    r"""^(
        [A-Z]\d{8}           |   # CIF tipo B26249896
        \d{8}[A-Z]           |   # NIF tipo 40313009N
        [XYZ]\d{7}[A-Z]          # NIE tipo X1234567A
    )$""",
    re.VERBOSE
)

def es_id_fiscal(token: str) -> bool:
    token = token.strip().replace(" ", "")
    return bool(PAT_ID_GENERAL.match(token))

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

# Para hoja NIE_Warnings (detección global en rango):
PAT_ID_GLOBAL = re.compile(
    r"""(
        [A-Z]\d{8}           |
        \d{8}[A-Z]           |
        [XYZ]\d{7}[A-Z]
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
    Flujo:
      - Recibe PDF + rango de páginas (1-based)
      - Agrupa texto por filas visuales (Y-cercano)
      - Para cada fila, ordena spans por X
      - Usa el primer token con pinta de NIF/CIF/NIE (9 chars-ish) como frontera:
          antes  -> nombre completo / razón social
          token  -> NIF/CIF/NIE
          después-> MARCA, MATRÍCULA
    """

    # 1. Cargar referencias / exclusiones (igual que antes)
    try:
        referencias_list: List[str] = json.loads(referencias)
        exclusiones_map: Dict[str, List[str]] = json.loads(exclusiones)
    except Exception:
        return {"error": "No se pudo leer referencias o exclusiones"}

    # 2. Ajustar páginas a base 0
    start_page = max(pagina_inicio - 1, 0)
    end_page = max(pagina_fin - 1, start_page)

    original_name = file.filename or "archivo.pdf"
    temp_pdf = f"/tmp/{uuid.uuid4().hex}_{original_name}"
    with open(temp_pdf, "wb") as f:
        f.write(await file.read())

    doc = fitz.open(temp_pdf)
    end_page = min(end_page, len(doc) - 1)

    # 3. Detectar IDs fiscales globales en el rango para la hoja de warnings
    ids_detectados = set()
    for pageno in range(start_page, end_page + 1):
        page_text = doc[pageno].get_text("text")
        for m in PAT_ID_GLOBAL.findall(page_text):
            ids_detectados.add(m.strip())

    # 4. Agrupar spans por fila visual (Y)
    TOL_Y = 4
    filas_totales: List[Dict[str, str]] = []

    # Bandera de "solo procesar después de ver una cabecera real"
    en_modo_tabla = False

    for pageno in range(start_page, end_page + 1):
        page = doc[pageno]
        page_dict = page.get_text("dict")

        # 4.1 Detectar si esta página contiene encabezados tipo
        # "NOMBRE Y APELLIDOS O RAZÓN SOCIAL", "NIF", "MARCA", "MATRÍCULA"
        pagina_tiene_header = False
        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    texto_span = span["text"].strip()
                    # si cualquier referencia coincide, asumimos que ya estamos en la tabla
                    for ref in referencias_list:
                        if header_match(texto_span, ref):
                            pagina_tiene_header = True
                            break

        if pagina_tiene_header:
            en_modo_tabla = True

        # Si aún no hemos visto headers, saltamos esta página
        if not en_modo_tabla:
            continue

        # 4.2 Construir filas "crudas": lista de (y_centro, spans_de_fila[])
        filas_pag_cruda: List[Tuple[float, List[dict]]] = []

        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                ys = [s["bbox"][1] for s in spans] + [s["bbox"][3] for s in spans]
                y_centro = (min(ys) + max(ys)) / 2.0

                # buscar fila existente cercana
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

                    # filtrar numeritos sueltos tipo número de página
                    if re.fullmatch(r"\d{1,3}", texto_span):
                        continue

                    # filtrar cabeceras repetidas
                    if any(header_match(texto_span, ref) for ref in referencias_list):
                        continue

                    # filtrar exclusiones explícitas
                    skip_it = False
                    for ref in referencias_list:
                        if texto_span in exclusiones_map.get(ref, []):
                            skip_it = True
                            break
                    if skip_it:
                        continue

                    x0, y0, x1, y1 = s["bbox"]
                    filas_pag_cruda[fila_idx][1].append({
                        "text": texto_span,
                        "x0": x0,
                        "y0": y0,
                        "x1": x1,
                        "y1": y1
                    })

        # 4.3 Parsear cada fila cruda usando el patrón NIF como ancla
        for (_, spans_de_fila) in filas_pag_cruda:
            if not spans_de_fila:
                continue

            # Ordenar los spans de la fila por posición izquierda→derecha
            spans_ordenados = sorted(spans_de_fila, key=lambda sp: sp["x0"])

            buffer_nombre: List[str] = []
            nif_val = ""
            marca_val = ""
            matric_val = ""

            fase = "nombre"

            for sp in spans_ordenados:
                token = sp["text"].strip()

                if fase == "nombre":
                    # ¿Este token es un posible NIF/CIF/NIE?
                    if es_id_fiscal(token):
                        nif_val = token
                        fase = "post-nif"
                    else:
                        buffer_nombre.append(token)

                elif fase == "post-nif":
                    # Después del NIF: primer token -> MARCA, segundo -> MATRÍCULA
                    if not marca_val:
                        marca_val = token
                    elif not matric_val:
                        matric_val = token
                    else:
                        # si hay tokens extra, los podemos concatenar a matrícula
                        matric_val = (matric_val + " " + token).strip()

            nombre_full = " ".join(buffer_nombre).strip()

            # descartamos filas totalmente vacías
            if not (nombre_full or nif_val or marca_val or matric_val):
                continue

            filas_totales.append({
                "NOMBRE": nombre_full,
                "NIF": nif_val,
                "MARCA": marca_val,
                "MATRICULA": matric_val
            })

    doc.close()

    # 5. Crear Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    ws.append([
        "NOMBRE Y APELLIDOS O RAZÓN SOCIAL",
        "NIF",
        "MARCA",
        "MATRÍCULA",
    ])

    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for fila in filas_totales:
        row_vals = [
            fila.get("NOMBRE", ""),
            fila.get("NIF", ""),
            fila.get("MARCA", ""),
            fila.get("MATRICULA", "")
        ]
        ws.append(row_vals)

        current_row_idx = ws.max_row
        nombre_val = ws.cell(current_row_idx, 1).value or ""
        if len(nombre_val) > 60:
            ws.cell(current_row_idx, 1).fill = amarillo

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
