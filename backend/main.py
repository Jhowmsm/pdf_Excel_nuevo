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
    title="PDF Table Extractor (RAW mode)",
    description="Extrae filas crudas de tablas PDF y las vuelca en Excel para post-procesar.",
    version="2.0.0"
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

# Detecta NIF/NIE/CIF
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

# Para la hoja NIE_Warnings (sacamos todos los IDs fiscales que aparezcan)
PAT_ID_GLOBAL = re.compile(
    r"""(
        [A-Z]\d{8}           |
        \d{8}[A-Z]           |
        [XYZ]\d{7}[A-Z]
    )""",
    re.VERBOSE
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

@app.post("/procesar/")
async def procesar_pdf(
    file: UploadFile = File(...),
    referencias: str = Form(...),
    exclusiones: str = Form("{}"),
    pagina_inicio: int = Form(...),
    pagina_fin: int = Form(...)
):
    """
    Modo RAW:
    - Lee únicamente las páginas [pagina_inicio-1 .. pagina_fin-1]
    - A partir de la primera página donde vea headers tipo 'NOMBRE...', 'NIF', etc.,
      empieza a capturar filas.
    - Para cada fila visual:
        - Ordena spans por X
        - Concatena TODOS los spans en orden -> RAW_LINE
        - Detecta el primer token que parece NIF/CIF/NIE -> NIF_DETECTADO
        - Después del NIF, primer token -> MARCA_POSIBLE
        - Después del NIF, segundo token -> MATRICULA_POSIBLE
    - Excel final: columnas RAW_LINE, NIF_DETECTADO, MARCA_POSIBLE, MATRICULA_POSIBLE
    """

    # 1. Parsear los extras del formulario
    try:
        referencias_list: List[str] = json.loads(referencias)
        exclusiones_map: Dict[str, List[str]] = json.loads(exclusiones)
    except Exception:
        return {"error": "No se pudo leer referencias o exclusiones"}

    start_page = max(pagina_inicio - 1, 0)
    end_page = max(pagina_fin - 1, start_page)

    # 2. Guardar PDF temporal
    original_name = file.filename or "archivo.pdf"
    temp_pdf = f"/tmp/{uuid.uuid4().hex}_{original_name}"
    with open(temp_pdf, "wb") as f:
        f.write(await file.read())

    # 3. Abrir con PyMuPDF
    doc = fitz.open(temp_pdf)
    end_page = min(end_page, len(doc) - 1)

    # 4. Detectar IDs fiscales globales (para la pestaña NIE_Warnings)
    ids_detectados = set()
    for pageno in range(start_page, end_page + 1):
        page_text = doc[pageno].get_text("text")
        for m in PAT_ID_GLOBAL.findall(page_text):
            ids_detectados.add(m.strip())

    TOL_Y = 4  # tolerancia vertical

    filas_totales = []

    en_modo_tabla = False  # aún no hemos visto cabecera de tabla

    # 5. Recorremos sólo el rango de páginas
    for pageno in range(start_page, end_page + 1):
        page = doc[pageno]
        page_dict = page.get_text("dict")

        # 5.1 ¿Esta página parece tener cabecera?
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

        # si todavía no hemos llegado a las tablas, saltar
        if not en_modo_tabla:
            continue

        # 5.2 Agrupar spans por fila visual Y
        filas_pag_cruda: List[Tuple[float, List[dict]]] = []

        for block in page_dict["blocks"]:
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                ys = [s["bbox"][1] for s in spans] + [s["bbox"][3] for s in spans]
                y_centro = (min(ys) + max(ys)) / 2.0

                # buscar fila con Y cercano
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

                    # descartamos cosas que no son contenido de fila:
                    # - numeritos tipo "51"
                    if re.fullmatch(r"\d{1,3}", texto_span):
                        continue

                    # - encabezados repetidos
                    if any(header_match(texto_span, ref) for ref in referencias_list):
                        continue

                    # - exclusiones explícitas (por ejemplo, el nombre del header)
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
                    })

        # 5.3 Procesar cada fila cruda en modo "lineal"
        for (_, spans_de_fila) in filas_pag_cruda:
            if not spans_de_fila:
                continue

            # ordenar TODOS los spans de la fila de izquierda a derecha
            spans_sorted = sorted(spans_de_fila, key=lambda sp: sp["x0"])

            # reconstrucción lineal
            tokens_line = [sp["text"] for sp in spans_sorted]
            raw_line = " ".join(tokens_line).strip()

            # detectar el primer identificador fiscal
            nif_detectado = ""
            marca_val = ""
            matric_val = ""

            fase = "antes_nif"
            for tok in tokens_line:
                if fase == "antes_nif":
                    if es_id_fiscal(tok):
                        nif_detectado = tok
                        fase = "despues_nif"
                elif fase == "despues_nif":
                    if not marca_val:
                        marca_val = tok
                    elif not matric_val:
                        matric_val = tok
                    else:
                        # Si aparece texto extra después de matrícula, podemos anexarlo a matrícula
                        matric_val = (matric_val + " " + tok).strip()

            # descartamos filas que están demasiado vacías
            if not raw_line:
                continue

            filas_totales.append({
                "RAW_LINE": raw_line,
                "NIF_DETECTADO": nif_detectado,
                "MARCA_POSIBLE": marca_val,
                "MATRICULA_POSIBLE": matric_val
            })

    doc.close()

    # 6. Crear Excel final en modo RAW
    wb = Workbook()
    ws = wb.active
    ws.title = "RAW_FILAS"

    ws.append([
        "RAW_LINE",
        "NIF_DETECTADO",
        "MARCA_POSIBLE",
        "MATRICULA_POSIBLE"
    ])

    for fila in filas_totales:
        ws.append([
            fila.get("RAW_LINE", ""),
            fila.get("NIF_DETECTADO", ""),
            fila.get("MARCA_POSIBLE", ""),
            fila.get("MATRICULA_POSIBLE", "")
        ])

    # Hoja secundaria con todos los IDs fiscales detectados en el rango
    if ids_detectados:
        ws_alerta = wb.create_sheet("NIE_Warnings")
        ws_alerta.append(["NIF/NIE/CIF Detectados"])
        for codigo in sorted(ids_detectados):
            ws_alerta.append([codigo])

    output_path = f"/tmp/resultado_{uuid.uuid4().hex}.xlsx"
    wb.save(output_path)

    return FileResponse(
        output_path,
        filename="resultado.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
