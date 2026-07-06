import os
import pandas as pd
import pdfplumber
from docx import Document

 # ====================== RESOLVE PATH ====================== #

def _resolve_path(path: str, workdir: str) -> str:
    if os.path.isabs(path):
        return path
    return os.path.join(workdir, path)

 # ====================== ANALYSE EXCEL ====================== #
def analyse_excel(path: str, sheet_name: str, workdir: str) -> dict:
    
    full_path = _resolve_path(path, workdir)

    if not os.path.exists(full_path):
        return {"status": "error", "message": f"File not found: {full_path}"}

    try:
        df = pd.read_excel(full_path, sheet_name=sheet_name)
    except Exception as e:
        return {"status": "error", "message": str(e)}

    if df.empty:
        return {"status": "error", "message": "Sheet is empty"}

    row_count = len(df)
    col_count = len(df.columns)
    columns = list(df.columns)

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    text_cols = df.select_dtypes(exclude="number").columns.tolist()

    stats = {}
    for col in numeric_cols:
        col_data = df[col].dropna()
        if col_data.empty:
            continue

        max_idx = col_data.idxmax()
        max_label = None
        if text_cols:
            max_label = str(df.loc[max_idx, text_cols[0]])

        stats[col] = {
          "min":     float(round(col_data.min(), 2)),
          "max":     float(round(col_data.max(), 2)),
          "average": float(round(col_data.mean(), 2)),
          "total":   float(round(col_data.sum(), 2)),
          "highest_in": max_label
}

    return {
        "status": "ok",
        "dataset_info": {
            "rows": row_count,
            "columns": col_count,
            "column_names": columns,
            "numeric_columns": numeric_cols,
            "text_columns": text_cols
        },
        "stats": stats
    }

 # ====================== SUMMARISE PDF ====================== #

def summarise_pdf(path: str, workdir: str) -> dict:
    
    full_path = _resolve_path(path, workdir)

    if not os.path.exists(full_path):
        return {"status": "error", "message": f"File not found: {full_path}"}

    pages_text = []
    with pdfplumber.open(full_path) as pdf:
        page_count = len(pdf.pages)
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                pages_text.append(text.strip())

    if not pages_text:
        return {"status": "error", "message": "No extractable text. PDF may be a scanned image."}

    full_text = "\n\n".join(pages_text)

    truncated = len(full_text) > 6000
    if truncated:
        full_text = full_text[:6000]

    return {
        "status": "ok",
        "page_count": page_count,
        "truncated": truncated,
        "content": full_text
    }

 # ====================== SUMMARISE WORD ====================== #

def summarise_word(path: str, workdir: str) -> dict:
   
    full_path = _resolve_path(path, workdir)

    if not os.path.exists(full_path):
        return {"status": "error", "message": f"File not found: {full_path}"}

    doc = Document(full_path)

    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    if not paragraphs:
        return {"status": "error", "message": "Document is empty"}

    full_text = "\n".join(paragraphs)

    truncated = len(full_text) > 6000
    if truncated:
        full_text = full_text[:6000]

    return {
        "status": "ok",
        "paragraph_count": len(paragraphs),
        "truncated": truncated,
        "content": full_text
    }
