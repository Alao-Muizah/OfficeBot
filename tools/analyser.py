import os
import pandas as pd
import pdfplumber
from docx import Document


def resolve_path(path: str, workdir: str) -> str:
    if os.path.isabs(path):
        return path
    return os.path.join(workdir, path)


# =============================================================================
# PUBLIC ENTRY POINT
# =============================================================================

def analyse_file(path: str, workdir: str, sheet_name: str = "Sheet1") -> dict:
    full_path = resolve_path(path, workdir)

    if not os.path.exists(full_path):
        return {"status": "error", "message": f"File not found: {full_path}"}

    ext = os.path.splitext(full_path)[1].lower()

    if ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
        return analyse_excel(full_path, sheet_name)
    elif ext == '.csv':
        return analyse_csv(full_path)
    elif ext == '.pdf':
        return summarise_pdf(full_path)
    elif ext == '.docx':
        return summarise_word(full_path)
    else:
        return {"status": "error", "message": f"Unsupported file type: '{ext}'. Supported: .xlsx, .csv, .pdf, .docx"}


# =============================================================================
# PRIVATE HELPERS
# =============================================================================

def analyse_excel(full_path: str, sheet_name: str) -> dict:
    try:
        df = pd.read_excel(full_path, sheet_name=sheet_name)
    except Exception as e:
        return {"status": "error", "message": str(e)}
    return analyse_dataframe(df)


def analyse_csv(path: str, workdir: str) -> dict:
    full_path = resolve_path(path, workdir)
    try:
        df = pd.read_csv(full_path)
    except Exception as e:
        return {"status": "error", "message": str(e)}
    return analyse_dataframe(df)


def analyse_dataframe(df: pd.DataFrame) -> dict:
    if df.empty:
        return {"status": "error", "message": "File is empty"}

    row_count = len(df)
    col_count = len(df.columns)
    columns = [str(c) for c in df.columns]
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    text_cols = df.select_dtypes(exclude="number").columns.tolist()

    stats = {}
    for col in numeric_cols:
        col_data = df[col].dropna()
        if col_data.empty:
            continue

        max_idx = col_data.idxmax()
        min_idx = col_data.idxmin()
        max_label = str(df.loc[max_idx, text_cols[0]]) if text_cols else None
        min_label = str(df.loc[min_idx, text_cols[0]]) if text_cols else None

        stats[str(col)] = {
            "min":        float(round(col_data.min(), 2)),
            "max":        float(round(col_data.max(), 2)),
            "average":    float(round(col_data.mean(), 2)),
            "median":     float(round(col_data.median(), 2)),
            "total":      float(round(col_data.sum(), 2)),
            "std_dev":    float(round(col_data.std(), 2)),
            "missing":    int(df[col].isna().sum()),
            "highest_in": max_label,
            "lowest_in":  min_label,
            "top_5":      [float(x) for x in col_data.nlargest(5).tolist()],
            "bottom_5":   [float(x) for x in col_data.nsmallest(5).tolist()]
        }

    return {
        "status":          "ok",
        "row_count":       int(row_count),
        "col_count":       int(col_count),
        "columns":         columns,
        "numeric_columns": [str(c) for c in numeric_cols],
        "text_columns":    [str(c) for c in text_cols],
        "missing_values":  {str(k): int(v) for k, v in df.isna().sum().items()},
        "sample_data":     df.head(5).astype(str).to_dict(orient="records"),
        "stats":           stats
    }


def summarise_pdf(path: str, workdir:str) -> dict:
    full_path = resolve_path(path, workdir)

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

    return {
        "status":     "ok",
        "file_type":  "pdf",
        "page_count": page_count,
        "truncated":  truncated,
        "content":    full_text[:6000] if truncated else full_text
    }


def summarise_word(path: str, workdir:str) -> dict:
    full_path = resolve_path(path, workdir)

    doc = Document(full_path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    if not paragraphs:
        return {"status": "error", "message": "Document is empty"}

    full_text = "\n".join(paragraphs)
    truncated = len(full_text) > 6000

    return {
        "status":          "ok",
        "file_type":       "word",
        "paragraph_count": len(paragraphs),
        "truncated":       truncated,
        "content":         full_text[:6000] if truncated else full_text
    }
