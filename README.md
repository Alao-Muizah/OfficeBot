# OfficeBot
An AI-powered CLI assistant for working with Word and Excel files using natural language instructions.

## What it does
- Create, read, summarise, and edit Word documents (.docx)
- Create, read, analyse, and edit Excel spreadsheets (.xlsx)
- Analyse Excel data and explain trends in plain English
- Summarise the contents of Word documents and PDFs
- Convert between formats: Word ↔ PDF, Excel ↔ CSV
- File operations: list, copy, rename, delete
- Works with files from any folder path on your computer
- Natural language interface — just describe what you want done

## Project structure
```
OfficeBot/
├── main.py              # CLI entry point
├── agent.py             # ReAct loop + Mistral API calls
├── state.py             # Session memory (open files, history)
├── config.py            # API key and model config
├── requirements.txt     # Dependencies
└── tools/
    ├── __init__.py
    ├── word.py          # Word document operations
    ├── excel.py         # Excel spreadsheet operations
    ├── analyser.py      # Excel analysis and PDF/Word summarisation
    ├── convert.py       # File format conversions
    └── filesystem.py    # File system operations
```

## Setup

**1. Clone the repo**
```bash
git clone https://github.com/Alao-Muizah/OfficeBot.git
cd OfficeBot
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Get a Mistral API key**

Go to [console.mistral.ai](https://console.mistral.ai) → API Keys → create a key.

**4. Set your API key**

Paste directly into `config.py`:
```python
API_KEY = "your_key_here"
```

**5. Run**
```bash
python main.py

# Custom working directory
python main.py --workdir "C:/Users/You/Documents"
```

## Example instructions

```
create a Word doc called report.docx with title "Q2 Report" and body "Sales summary."
add a paragraph to report.docx that says "Revenue grew by 18% this quarter."
insert a table into report.docx with headers Name, Region, Revenue and 3 rows of sample data
replace "Q2" with "Second Quarter" in report.docx
convert report.docx to PDF and save it as report.pdf
summarise report.docx
summarise C:\Users\You\Documents\contract.pdf

create an excel file called sales.xlsx with a sheet named Data
add headers Name, Units, Revenue to sales.xlsx sheet Data
add 5 rows to sales.xlsx sheet Data starting at row 2, rows: Alice 120 50000, Bob 95 38000, Carol 200 82000, Dave 150 61000, Eve 180 74000
apply a SUM formula in cell C6 of sales.xlsx sheet Data
analyse sales.xlsx sheet Data
convert sales.xlsx sheet Data to CSV and save as sales.csv

list all my files
rename report.docx to report_final.docx
copy sales.xlsx to sales_backup.xlsx
```

## How it works

OfficeBot uses a **ReAct loop** (Reason + Act):

1. User types an instruction
2. Instruction is sent to the Mistral LLM along with tool definitions
3. LLM picks the right tool and arguments
4. The tool (Python function) runs and returns a result
5. Result is fed back to the LLM
6. LLM either calls another tool or gives the final response

Session state (open files, last action, conversation history) persists across turns within a session so the agent remembers context.

## Dependencies
- `Python` — Overall language
- `mistralai` — LLM API
- `python-docx` — Word document operations
- `openpyxl` — Excel operations
- `pandas` — Excel data analysis
- `pdfplumber` — PDF text extraction
- `reportlab` — PDF generation
- `rich` — CLI formatting
