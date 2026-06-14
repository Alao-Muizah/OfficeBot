import json
import os
import time
from mistralai import Mistral
from config import API_KEY, MODEL
from state import SessionState

from tools.word import (
    create_word_doc, read_word_doc, append_paragraph,
    add_heading, replace_text, set_font, insert_table, text_format
)
from tools.excel import (
    create_excel, read_sheet, write_headers, write_rows,
    write_cell, apply_formula, format_cell, set_column_width,
    add_sheet, list_sheets
)
from tools.converter import (
    convert_word_to_pdf, convert_pdf_to_word,
    convert_excel_to_csv, convert_csv_to_excel
)
from tools.filesystem import (
    fetch_file, list_files, copy_file, delete_file, rename_file
)
from tools.analyser import analyse_excel, summarise_pdf, summarise_word

client = Mistral(api_key=API_KEY)


# =============================================================================
# TOOL DEFINITIONS
# =============================================================================

TOOLS = [
     # ====================== WORD ====================== #
    {
        "type": "function",
        "function": {
            "name": "create_word_doc",
            "description": "Creates a new Word (.docx) document with a title and body text.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":  {"type": "string", "description": "Filename e.g. 'report.docx'"},
                    "title": {"type": "string", "description": "Document title, added as Heading 1"},
                    "body":  {"type": "string", "description": "Main body text"}
                },
                "required": ["path", "title", "body"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_word_doc",
            "description": "Reads and returns all text from an existing Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Filename or full path of the document"}
                },
                "required": ["path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "append_paragraph",
            "description": "Adds a new paragraph at the end of an existing Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the document"},
                    "text": {"type": "string", "description": "Paragraph text to add"}
                },
                "required": ["path", "text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "add_heading",
            "description": "Adds a heading to an existing Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":  {"type": "string",  "description": "Path to the document"},
                    "text":  {"type": "string",  "description": "Heading text"},
                    "level": {"type": "integer", "description": "Heading level 1-9, 1 is biggest"}
                },
                "required": ["path", "text", "level"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "replace_text",
            "description": "Replaces all occurrences of a word or phrase in a Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":     {"type": "string", "description": "Path to the document"},
                    "old_text": {"type": "string", "description": "Text to find"},
                    "new_text": {"type": "string", "description": "Replacement text"}
                },
                "required": ["path", "old_text", "new_text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "set_font",
            "description": "Sets the font name and size for all text in a Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":      {"type": "string",  "description": "Path to the document"},
                    "font_name": {"type": "string",  "description": "Font name e.g. 'Arial'"},
                    "font_size": {"type": "integer", "description": "Font size in points e.g. 12"}
                },
                "required": ["path", "font_name", "font_size"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "insert_table",
            "description": "Inserts a table with headers and data rows at the end of a Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":    {"type": "string", "description": "Path to the document"},
                    "headers": {"type": "array", "items": {"type": "string"}, "description": "Column header names"},
                    "rows":    {"type": "array", "items": {"type": "array"}, "description": "Rows of data"}
                },
                "required": ["path", "headers", "rows"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "bold_text",
            "description": "Makes a specific word or phrase bold in a Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":        {"type": "string", "description": "Path to the document"},
                    "target_text": {"type": "string", "description": "The text to bold"}
                },
                "required": ["path", "target_text"]
            }
        }
    },

    # ====================== EXCEL ====================== #
    {
        "type": "function",
        "function": {
            "name": "create_excel",
            "description": "Creates a new Excel (.xlsx) workbook with one sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Filename e.g. 'sales.xlsx'"},
                    "sheet_name": {"type": "string", "description": "Name for the first sheet"}
                },
                "required": ["path", "sheet_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_sheet",
            "description": "Reads all data from an Excel sheet, returns as list of row dicts.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string", "description": "Sheet name to read"}
                },
                "required": ["path", "sheet_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "write_headers",
            "description": "Writes column headers into row 1, styled bold with gray background.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "headers":    {"type": "array", "items": {"type": "string"}, "description": "Header names"}
                },
                "required": ["path", "sheet_name", "headers"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "write_rows",
            "description": "Writes multiple rows of data into a sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string",  "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string",  "description": "Sheet name"},
                    "rows":       {"type": "array",   "items": {"type": "array"}, "description": "Rows of data"},
                    "start_row":  {"type": "integer", "description": "Row number to start at. Use 2 if headers are in row 1."}
                },
                "required": ["path", "sheet_name", "rows", "start_row"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "write_cell",
            "description": "Writes a single value to a specific cell e.g. B3.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "cell_ref":   {"type": "string", "description": "Cell reference e.g. 'B3'"},
                    "value":      {"description": "Value to write"}
                },
                "required": ["path", "sheet_name", "cell_ref", "value"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "apply_formula",
            "description": "Writes an Excel formula into a cell e.g. =SUM(B2:B10).",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "cell_ref":   {"type": "string", "description": "Cell to put formula in"},
                    "formula":    {"type": "string", "description": "Formula string, must start with ="}
                },
                "required": ["path", "sheet_name", "cell_ref", "formula"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "format_cell",
            "description": "Applies formatting to a single cell: bold, font size, font color.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string",  "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string",  "description": "Sheet name"},
                    "cell_ref":   {"type": "string",  "description": "Cell reference e.g. 'A1'"},
                    "bold":       {"type": "boolean", "description": "Make text bold"},
                    "font_size":  {"type": "integer", "description": "Font size in points"},
                    "font_color": {"type": "string",  "description": "Hex color without # e.g. FF0000 for red"}
                },
                "required": ["path", "sheet_name", "cell_ref", "bold", "font_size", "font_color"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "set_column_width",
            "description": "Sets the width of a column in an Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string", "description": "Sheet name"},
                    "column":     {"type": "string", "description": "Column letter e.g. 'A'"},
                    "width":      {"type": "number", "description": "Width. 15 is default."}
                },
                "required": ["path", "sheet_name", "column", "width"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "add_sheet",
            "description": "Adds a new empty sheet to an existing Excel workbook.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string", "description": "Name for the new sheet"}
                },
                "required": ["path", "sheet_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_sheets",
            "description": "Returns the names of all sheets in an Excel workbook.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel file"}
                },
                "required": ["path"]
            }
        }
    },

    # ====================== CONVERTER ====================== #
    {
        "type": "function",
        "function": {
            "name": "convert_word_to_pdf",
            "description": "Converts a Word (.docx) file to PDF.",
            "parameters": {
                "type": "object",
                "properties": {
                    "input_path":  {"type": "string", "description": "Path to the source .docx file"},
                    "output_path": {"type": "string", "description": "Path for the output .pdf file"}
                },
                "required": ["input_path", "output_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "convert_pdf_to_word",
            "description": "Extracts text from a PDF and writes it into a Word document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "input_path":  {"type": "string", "description": "Path to the source .pdf file"},
                    "output_path": {"type": "string", "description": "Path for the output .docx file"}
                },
                "required": ["input_path", "output_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "convert_excel_to_csv",
            "description": "Converts one sheet from an Excel file to a CSV file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "input_path":  {"type": "string", "description": "Path to the source .xlsx file"},
                    "output_path": {"type": "string", "description": "Path for the output .csv file"},
                    "sheet_name":  {"type": "string", "description": "Sheet name to convert"}
                },
                "required": ["input_path", "output_path", "sheet_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "convert_csv_to_excel",
            "description": "Converts a CSV file into an Excel workbook.",
            "parameters": {
                "type": "object",
                "properties": {
                    "input_path":  {"type": "string", "description": "Path to the source .csv file"},
                    "output_path": {"type": "string", "description": "Path for the output .xlsx file"},
                    "sheet_name":  {"type": "string", "description": "Sheet name in the new workbook"}
                },
                "required": ["input_path", "output_path", "sheet_name"]
            }
        }
    },

    # ====================== FILESYSTEM ====================== #
    {
        "type": "function",
        "function": {
            "name": "fetch_file",
            "description": "Checks a file exists and returns its name, size, type, and last modified date.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Filename or full path to check"}
                },
                "required": ["path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_files",
            "description": "Lists all Word, Excel, CSV, and PDF files in a directory.",
            "parameters": {
                "type": "object",
                "properties": {
                    "directory": {"type": "string", "description": "Directory to list. Use '.' for the working directory."}
                },
                "required": ["directory"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "copy_file",
            "description": "Copies a file to a new location or filename.",
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path":      {"type": "string", "description": "Path of the file to copy"},
                    "destination_path": {"type": "string", "description": "Destination path or filename"}
                },
                "required": ["source_path", "destination_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "delete_file",
            "description": "Permanently deletes a file. Only works within the working directory.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Filename or path of the file to delete"}
                },
                "required": ["path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "rename_file",
            "description": "Renames a file within the same directory.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":     {"type": "string", "description": "Current filename or path"},
                    "new_name": {"type": "string", "description": "New filename e.g. 'report_v2.docx'"}
                },
                "required": ["path", "new_name"]
            }
        }
    },

    # ====================== ANALYSIS & SUMMARISATION ====================== #
    {
        "type": "function",
        "function": {
            "name": "analyse_excel",
            "description": "Analyses an Excel sheet using pandas and returns statistics — min, max, average, total per numeric column. Use this when the user asks to summarise, analyse, or explain trends in an Excel file.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path":       {"type": "string", "description": "Path to the Excel file"},
                    "sheet_name": {"type": "string", "description": "Sheet name to analyse"}
                },
                "required": ["path", "sheet_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "summarise_pdf",
            "description": "Extracts text from a PDF file so you can summarise what it contains. Use when the user asks what a PDF is about or wants a summary of it.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the PDF file"}
                },
                "required": ["path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "summarise_word",
            "description": "Reads a Word document and returns its content so you can summarise what it contains. Use when the user asks what a Word document is about or wants a summary of it.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Word document"}
                },
                "required": ["path"]
            }
        }
    }
]


# =============================================================================
# DISPATCH
# =============================================================================

def dispatch_tool(tool_name: str, tool_args: dict, workdir: str) -> str:
    if   tool_name == "create_word_doc":      result = create_word_doc(**tool_args, workdir=workdir)
    elif tool_name == "read_word_doc":         result = read_word_doc(**tool_args, workdir=workdir)
    elif tool_name == "append_paragraph":      result = append_paragraph(**tool_args, workdir=workdir)
    elif tool_name == "add_heading":           result = add_heading(**tool_args, workdir=workdir)
    elif tool_name == "replace_text":          result = replace_text(**tool_args, workdir=workdir)
    elif tool_name == "set_font":              result = set_font(**tool_args, workdir=workdir)
    elif tool_name == "insert_table":          result = insert_table(**tool_args, workdir=workdir)
    elif tool_name == "bold_text":             result = text_format(**tool_args, workdir=workdir)
    elif tool_name == "create_excel":          result = create_excel(**tool_args, workdir=workdir)
    elif tool_name == "read_sheet":            result = read_sheet(**tool_args, workdir=workdir)
    elif tool_name == "write_headers":         result = write_headers(**tool_args, workdir=workdir)
    elif tool_name == "write_rows":            result = write_rows(**tool_args, workdir=workdir)
    elif tool_name == "write_cell":            result = write_cell(**tool_args, workdir=workdir)
    elif tool_name == "apply_formula":         result = apply_formula(**tool_args, workdir=workdir)
    elif tool_name == "format_cell":           result = format_cell(**tool_args, workdir=workdir)
    elif tool_name == "set_column_width":      result = set_column_width(**tool_args, workdir=workdir)
    elif tool_name == "add_sheet":             result = add_sheet(**tool_args, workdir=workdir)
    elif tool_name == "list_sheets":           result = list_sheets(**tool_args, workdir=workdir)
    elif tool_name == "convert_word_to_pdf":   result = convert_word_to_pdf(**tool_args, workdir=workdir)
    elif tool_name == "convert_pdf_to_word":   result = convert_pdf_to_word(**tool_args, workdir=workdir)
    elif tool_name == "convert_excel_to_csv":  result = convert_excel_to_csv(**tool_args, workdir=workdir)
    elif tool_name == "convert_csv_to_excel":  result = convert_csv_to_excel(**tool_args, workdir=workdir)
    elif tool_name == "fetch_file":            result = fetch_file(**tool_args, workdir=workdir)
    elif tool_name == "list_files":            result = list_files(**tool_args, workdir=workdir)
    elif tool_name == "copy_file":             result = copy_file(**tool_args, workdir=workdir)
    elif tool_name == "delete_file":           result = delete_file(**tool_args, workdir=workdir)
    elif tool_name == "rename_file":           result = rename_file(**tool_args, workdir=workdir)
    elif tool_name == "analyse_excel":         result = analyse_excel(**tool_args, workdir=workdir)
    elif tool_name == "summarise_pdf":         result = summarise_pdf(**tool_args, workdir=workdir)
    elif tool_name == "summarise_word":        result = summarise_word(**tool_args, workdir=workdir)
    else:
        result = {"status": "error", "message": f"Unknown tool: {tool_name}"}

    return json.dumps(result, indent=2)


# =============================================================================
# REACT LOOP
# =============================================================================

def run_agent(user_message: str, state: SessionState) -> str:

    workdir = state.workdir

    system_prompt = f"""You are OfficeBot. Your only job is to execute the user's instruction using the correct tool.

    Working directory: {workdir}

    {state.get_context_summary()}

    - Execute the user's instruction exactly as stated. Use the exact filenames, titles, values, and data the user provides. Do not invent or substitute anything.
    - Call only the tool the instruction requires. Do not call extra tools before or after.
    - Use only the filename in path arguments unless the user gives a full path — in that case use the full path exactly.
    - If the instruction is complete, act immediately without asking questions.
    - After calling a tool, respond in plain English explaining what was done or what the result means. Never show raw JSON or numbers without explanation.
    - For analyse_excel: after getting the stats, narrate the trends clearly — explain what the dataset is about, which values are high or low, and what patterns stand out.
    - For summarise_pdf and summarise_word: after getting the content, give a clear plain English summary of what the document is about. Do not repeat the raw text.
    """

    if len(state.history) > 6:
        state.history = state.history[-6:]

    state.history.append({"role": "user", "content": user_message})

    messages = (
        [{"role": "system", "content": system_prompt}]
        + state.history
    )

    while True:

        for attempt in range(3):
            try:
                response = client.chat.complete(
                    model=MODEL,
                    messages=messages,
                    tools=TOOLS,
                    tool_choice="auto",
                    max_tokens=1024
                )
                break
            except Exception as e:
                if attempt < 2:
                    print(f"  [retrying {attempt + 1}/3...]")
                    time.sleep(2)
                    continue
                return f"Connection error: {str(e)}"

        message = response.choices[0].message
        finish_reason = str(response.choices[0].finish_reason)

        if "tool_calls" in finish_reason and message.tool_calls:

            assistant_msg = {
                "role": "assistant",
                "content": message.content or "",
                "tool_calls": [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {
                            "name": tc.function.name,
                            "arguments": tc.function.arguments
                        }
                    }
                    for tc in message.tool_calls
                ]
            }
            messages.append(assistant_msg)
            state.history.append(assistant_msg)

            for tool_call in message.tool_calls:
                tool_name = tool_call.function.name
                tool_args = json.loads(tool_call.function.arguments)

                print(f"   [tool: {tool_name} {tool_args}]")

                tool_result_str = dispatch_tool(tool_name, tool_args, workdir)
                tool_result = json.loads(tool_result_str)

                if tool_result.get("status") == "ok":
                    state.set_last_action(tool_result.get("message", tool_name))

                    tracked_path = (
                        tool_result.get("path")
                        or tool_result.get("output")
                        or tool_args.get("path")
                        or tool_args.get("input_path")
                    )

                    if tracked_path:
                        filename = os.path.basename(tracked_path)
                        state.register_file(
                            filename=filename,
                            full_path=tracked_path,
                            note=tool_result.get("message", "used")
                        )

                    if tool_name == "delete_file":
                        state.remove_file(os.path.basename(tool_args.get("path", "")))
                    elif tool_name == "rename_file":
                        state.remove_file(os.path.basename(tool_args.get("path", "")))

                tool_msg = {
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": tool_result_str
                }
                messages.append(tool_msg)
                state.history.append(tool_msg)

        else:
            final_msg = {"role": "assistant", "content": message.content or ""}
            state.history.append(final_msg)
            return message.content
        
