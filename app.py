
"""
Economic Data Summarizer (Desktop, CustomTkinter + Pandas + OpenAI Responses API)

What it does
------------
A native-style desktop app that:
1) Loads an Excel sheet (headers + last 24 rows).
2) Runs a TWO-STEP LLM pipeline using the OpenAI *Responses API*:
   - Step 1: Extract facts for every column, infer units (using example text if provided).
   - Step 2: Generate the final paragraph, optionally mimicking the example text with a
             Strict / Loose / None style setting.
3) Optionally edits the final paragraph with a separate LLM pass based on user requirements.
4) Optionally captures the target economy (user input or inferred from example text).

Key requirements satisfied
--------------------------
- UI: customtkinter
- Excel processing: pandas
- OpenAI integration: official `openai` Python library (v1+), Responses API (NOT chat.completions)
- Default model: gpt-5.2
- Reasoning effort: high
- Verbosity: high
- Non-blocking UI: background thread + queue polling
- Basic error handling: popups for missing key/file/sheet/API failures
- No hardcoded API keys

Notes
-----
- This app assumes the sheet is already reasonably clean (no big blank blocks, missing values, etc.).
  A visible warning is shown in the UI as requested.
- For .xlsx: openpyxl is used via pandas.
- For legacy .xls: xlrd is required and may still fail for some modern-encrypted/odd XLS files.

Run
---
python economic_data_summarizer.py
"""

from __future__ import annotations

import json
import os
import queue
import threading
import traceback
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

import customtkinter as ctk
from tkinter import filedialog, messagebox

from openai import OpenAI


# ----------------------------
# LLM prompt + schema helpers
# ----------------------------

STEP1_JSON_SCHEMA: Dict[str, Any] = {
    "name": "economic_data_facts",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "economy": {"type": "string"},
            "economy_source": {
                "type": "string",
                "enum": ["user", "example_text", "inferred", "unknown"],
            },
            "economy_confidence": {"type": "string", "enum": ["high", "medium", "low"]},
            "dataset_overview": {"type": "string"},
            "variables": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "name": {"type": "string"},
                        "unit": {"type": "string"},
                        "unit_confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                        "example_mentions": {"type": "boolean"},
                        "latest": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "period": {"type": "string"},
                                "value": {"type": "string"},
                            },
                            "required": ["period", "value"],
                        },
                        "previous": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "period": {"type": "string"},
                                "value": {"type": "string"},
                            },
                            "required": ["period", "value"],
                        },
                        "trend_summary": {"type": "string"},
                        "comparison_to_trend": {"type": "string"},
                        "comparison_to_previous": {"type": "string"},
                        "data_quality_notes": {"type": "string"},
                    },
                    "required": [
                        "name",
                        "unit",
                        "unit_confidence",
                        "example_mentions",
                        "latest",
                        "previous",
                        "trend_summary",
                        "comparison_to_trend",
                        "comparison_to_previous",
                        "data_quality_notes",
                    ],
                },
            },
        },
        "required": [
            "economy",
            "economy_source",
            "economy_confidence",
            "dataset_overview",
            "variables",
        ],
    },
}


def df_to_markdown_table(df: pd.DataFrame, max_cols: int = 50) -> Tuple[str, Optional[str]]:
    """
    Convert a dataframe to a markdown table (string) while keeping values as raw as possible.

    Returns: (markdown_table, warning_or_none)
    """
    warning = None
    df2 = df.copy()

    # Round numeric data to 1 decimal place for all columns except the first (time).
    # This happens before converting values to strings for LLM input.
    if df2.shape[1] > 1:
        for col in df2.columns[1:]:
            series = df2[col]
            if pd.api.types.is_numeric_dtype(series):
                df2[col] = series.round(1)
            else:
                numeric = pd.to_numeric(series, errors="coerce")
                if numeric.notna().any():
                    rounded = numeric.round(1)
                    df2[col] = series.where(numeric.isna(), rounded)

    # Convert everything to string-ish for safety (avoid "1.0" vs "1" surprises)
    # Keep NaN as empty string.
    df2 = df2.where(pd.notnull(df2), "")
    df2 = df2.astype(str)

    if df2.shape[1] > max_cols:
        warning = (
            f"Input has {df2.shape[1]} columns; truncating to the first {max_cols} columns to "
            "stay within prompt limits."
        )
        df2 = df2.iloc[:, :max_cols]

    try:
        md = df2.to_markdown(index=False)
    except Exception:
        # Fallback if tabulate isn't available or to_markdown fails for any reason.
        headers = list(df2.columns)
        rows = df2.values.tolist()
        md_lines = []
        md_lines.append("| " + " | ".join(headers) + " |")
        md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
        for r in rows:
            md_lines.append("| " + " | ".join(str(x) for x in r) + " |")
        md = "\n".join(md_lines)

    return md, warning


def build_step1_input(markdown_table: str, example_text: str, economy_input: str) -> str:
    """
    Step 1: Data extraction + unit deduction + broad analysis for every column.
    """
    example_text = (example_text or "").strip()
    economy_input = (economy_input or "").strip()

    return f"""
You are an economic data analyst. You will be given a table extracted from Excel containing headers (variable names)
and the last ~24 observations. Units are often NOT stated in the sheet.

Your job:
- Analyze EVERY variable/column in the table.
- Infer the unit for each variable:
  - If Example Text is provided, deduce units and conventions by cross-referencing phrasing (e.g., "% YOY", "YTD % YOY" "QOQ SAAR",
    "index", "bps", "level", "USD", "CNY", etc.).
  - If Example Text is empty, infer the most likely standard economic unit from the variable name and values.
- Identify for each variable:
  - Latest reading (use the last non-empty row for that column)
  - Previous reading (use the prior non-empty row)
  - Recent trend over the last ~24 rows (direction, volatility, turning points)
  - How the latest compares with the trend and with the previous reading
- CRITICAL: Numerical fidelity
  - Copy values EXACTLY as they appear in the table cells (including minus signs, decimals, commas).
  - Do not invent missing numbers. If something is missing/blank, describe that in data_quality_notes.

Also determine whether each variable is mentioned in the Example Text (example_mentions = true/false). If Example Text
is empty, set example_mentions=false for all variables.

ECONOMY CONTEXT:
- If Economy Input is provided (non-empty), set:
  - economy = Economy Input
  - economy_source = "user"
  - economy_confidence = "high"
- Else if the Example Text clearly specifies the economy, set economy accordingly and economy_source="example_text".
- Else if you can infer the economy from variable names or context, set economy and economy_source="inferred".
- Otherwise set economy="" and economy_source="unknown".
- Set economy_confidence based on your evidence (high/medium/low).

Return ONLY valid JSON that conforms EXACTLY to the provided JSON schema.

Economy Input (may be empty):
<<<ECONOMY
{economy_input}
ECONOMY>>>

Example Text (may be empty):
<<<EXAMPLE_TEXT
{example_text}
EXAMPLE_TEXT>>>

Excel slice (headers + last ~24 rows) as markdown:
<<<TABLE
{markdown_table}
TABLE>>>
""".strip()


def build_step2_input(step1_facts_json: str, example_text: str, strictness: str, economy_input: str) -> str:
    """
    Step 2: Final paragraph generation + variable filtering + style control.
    """
    example_text = (example_text or "").strip()
    strictness = (strictness or "None").strip()
    economy_input = (economy_input or "").strip()

    # Prompt logic requested in PRD:
    # - Strict / Loose: only write about variables present in the Example Text (and present in Step1 facts)
    # - None: cover all variables analyzed in Step1
    # Special-case: if no example text provided, ignore Strict/Loose filtering and write a standard summary.
    return f"""
You are an economist writing standardized data paragraphs for a report.

You are given:
1) Structured facts for each variable (JSON) from a prior extraction step.
2) An optional Example Text (may be empty).
3) A Style Strictness setting: Strict / Loose / None.

CRUCIAL FILTERING:
- If Example Text is NOT empty AND Style Strictness is Strict or Loose:
  - Only write about variables that are (a) present in the facts JSON AND (b) mentioned in the Example Text.
  - Omit all other variables entirely.
- If Style Strictness is None:
  - Write about ALL variables in the facts JSON.
- If Example Text is empty:
  - Write a standard professional summary covering ALL variables in the facts JSON, regardless of strictness.

STYLE:
- Strict:
  - Use the exact sentence structure, vocabulary, and formatting of the provided example.
  - Only swap out numbers and directional words (e.g., "rose" vs "fell") to fit the new data.
  - Match the number of paragraphs in the Example Text
- Loose:
  - Adopt the general professional tone and length of the example, but you may restructure naturally.
  - Match the number of paragraphs in the Example Text
- None:
  - Write a standard, professional economic data summary focusing on clarity and conciseness.
  - Choose one or multiple paragraphs as appropriate.


NUMERICAL ACCURACY:
- Use the units and exact numbers as provided in the facts JSON.
- Do not compute new values unless they are trivial restatements explicitly supported by the facts.

ECONOMY CONTEXT:
- If Economy Input is provided (non-empty), treat it as authoritative and reflect it in the summary when relevant.
- If Economy Input is empty, you may infer the economy from Example Text or the facts JSON, but do not invent one.

OUTPUT:
- Return ONLY the final text (no title, no bullets, no JSON).
- DO NOT describe the structure and scope of the dataset 

Style Strictness: {strictness}

Economy Input (may be empty):
<<<ECONOMY
{economy_input}
ECONOMY>>>

Example Text (may be empty):
<<<EXAMPLE_TEXT
{example_text}
EXAMPLE_TEXT>>>

Facts JSON:
<<<FACTS_JSON
{step1_facts_json}
FACTS_JSON>>>
""".strip()


def build_edit_input(original_summary: str, edit_requirements: str) -> str:
    """
    Optional post-edit step: improve logic and language while following user requirements.
    """
    original_summary = (original_summary or "").strip()
    edit_requirements = (edit_requirements or "").strip()

    return f"""
You are an economist and editor. Your task is to check the logic and improve the clarity and language of the original summary.
Follow the user's editing requirements if provided.

Editing requirements (may be empty):
<<<REQUIREMENTS
{edit_requirements}
REQUIREMENTS>>>

Original summary:
<<<SUMMARY
{original_summary}
SUMMARY>>>

Rules:
- Preserve the meaning and all numbers/units exactly as written, unless there is an obvious typo.
- Do not add new facts that are not present in the original summary.
- Keep the same number of paragraphs and structure, unless the user provides alternative editing requirements.
- Output only text with no title, no bullets, and no JSON.
""".strip()


def safe_get_output_text(resp: Any) -> str:
    """
    The OpenAI Python SDK typically provides resp.output_text for Responses API.
    This helper keeps things resilient across minor SDK variations.
    """
    if hasattr(resp, "output_text") and isinstance(resp.output_text, str):
        return resp.output_text
    # Fallback: attempt to reconstruct from output items
    try:
        chunks = []
        for item in getattr(resp, "output", []) or []:
            for c in getattr(item, "content", []) or []:
                t = getattr(c, "text", None)
                if isinstance(t, str):
                    chunks.append(t)
        return "\n".join(chunks).strip()
    except Exception:
        return str(resp)


# ----------------------------
# Excel loading
# ----------------------------

@dataclass
class ExcelSlice:
    markdown_table: str
    warning: Optional[str]
    columns: List[str]
    row_count: int


def load_excel_slice(path: str, sheet_name: str, last_n_rows: int = 24) -> ExcelSlice:
    """
    Read the Excel file, get headers and last N rows, return markdown table.

    Assumptions:
    - Row 1 in Excel contains variable names (headers).
    - The sheet is mostly rectangular.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")

    try:
        # dtype=object keeps "raw-ish" values (numbers remain numbers; strings remain strings; dates can be mixed).
        df = pd.read_excel(path, sheet_name=sheet_name, header=0, dtype=object)
    except ValueError as e:
        # Usually sheet not found; show available sheets if possible
        try:
            xl = pd.ExcelFile(path)
            available = xl.sheet_names
            raise ValueError(
                f"Could not read sheet '{sheet_name}'. Available sheets: {available}"
            ) from e
        except Exception:
            raise
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file. Details: {e}") from e

    if df is None or df.empty:
        raise ValueError("The selected sheet is empty (no data found).")

    # Drop fully empty rows (helps with trailing blank blocks)
    df = df.dropna(axis=0, how="all")

    if df.empty:
        raise ValueError("After removing fully blank rows, no data remains in the selected sheet.")

    # Take last N rows
    df_tail = df.tail(last_n_rows)

    md, warning = df_to_markdown_table(df_tail, max_cols=50)
    return ExcelSlice(
        markdown_table=md,
        warning=warning,
        columns=[str(c) for c in df.columns.tolist()],
        row_count=df_tail.shape[0],
    )


# ----------------------------
# OpenAI pipeline
# ----------------------------

class PipelineError(RuntimeError):
    pass


def run_two_step_pipeline(
    api_key: str,
    model: str,
    markdown_table: str,
    example_text: str,
    strictness: str,
    economy_input: str,
) -> Tuple[str, str]:
    """
    Returns: (step1_facts_json_pretty, final_paragraph)
    """
    client = OpenAI(api_key=api_key)

    # Step 1 (Structured JSON)
    step1_input = build_step1_input(
        markdown_table=markdown_table,
        example_text=example_text,
        economy_input=economy_input,
    )

    try:
        step1_resp = client.responses.create(
            model=model,
            input=step1_input,
            reasoning={"effort": "high"},
            text={
                "verbosity": "high",
                "format": {
                    "type": "json_schema",
                    **STEP1_JSON_SCHEMA,  # includes name/strict/schema
                },
            },
        )
    except Exception as e:
        raise PipelineError(f"OpenAI API call failed in Step 1. Details: {e}") from e

    step1_text = safe_get_output_text(step1_resp)
    if not step1_text:
        raise PipelineError("Step 1 returned empty output.")

    try:
        step1_obj = json.loads(step1_text)
    except json.JSONDecodeError as e:
        # Even with schema, be defensive and show a snippet
        snippet = step1_text[:1500]
        raise PipelineError(
            "Step 1 output was not valid JSON (unexpected). "
            f"JSON error: {e}. Output snippet:\n{snippet}"
        ) from e

    step1_pretty = json.dumps(step1_obj, indent=2, ensure_ascii=False)

    # Step 2 (Final paragraph)
    step2_input = build_step2_input(
        step1_facts_json=step1_pretty,
        example_text=example_text,
        strictness=strictness,
        economy_input=economy_input,
    )

    try:
        step2_resp = client.responses.create(
            model=model,
            input=step2_input,
            reasoning={"effort": "high"},
            text={"verbosity": "high"},
        )
    except Exception as e:
        raise PipelineError(f"OpenAI API call failed in Step 2. Details: {e}") from e

    final_paragraph = safe_get_output_text(step2_resp).strip()
    if not final_paragraph:
        raise PipelineError("Step 2 returned empty output.")

    return step1_pretty, final_paragraph


def run_edit_pipeline(
    api_key: str,
    model: str,
    original_summary: str,
    edit_requirements: str,
) -> str:
    """
    Returns: edited_paragraph
    """
    client = OpenAI(api_key=api_key)
    edit_input = build_edit_input(original_summary=original_summary, edit_requirements=edit_requirements)

    try:
        edit_resp = client.responses.create(
            model=model,
            input=edit_input,
            reasoning={"effort": "high"},
            text={"verbosity": "high"},
        )
    except Exception as e:
        raise PipelineError(f"OpenAI API call failed in Edit step. Details: {e}") from e

    edited_paragraph = safe_get_output_text(edit_resp).strip()
    if not edited_paragraph:
        raise PipelineError("Edit step returned empty output.")

    return edited_paragraph


# ----------------------------
# UI (CustomTkinter)
# ----------------------------

class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.title("AI Workflow: Economic Data Summarizer (Tony Lo, March 2026)")
        self.geometry("1100x780")
        self.minsize(980, 700)

        self._result_queue: "queue.Queue[Tuple[str, str, Optional[str]]]" = queue.Queue()
        self._worker_thread: Optional[threading.Thread] = None
        self._placeholder_text = "Your generated paragraph will appear here."

        # Layout: left inputs, right output
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=1)

        self._build_left_panel()
        self._build_right_panel()

        self.after(120, self._poll_queue)

    def _build_left_panel(self) -> None:
        self.left = ctk.CTkFrame(self)
        self.left.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        self.left.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(self.left, text="Inputs", font=ctk.CTkFont(size=18, weight="bold"))
        title.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))

        # File selection
        file_frame = ctk.CTkFrame(self.left)
        file_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=8)
        file_frame.grid_columnconfigure(0, weight=1)

        self.file_path_var = ctk.StringVar(value="")
        self.file_entry = ctk.CTkEntry(file_frame, textvariable=self.file_path_var)
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(10, 8), pady=10)

        browse_btn = ctk.CTkButton(file_frame, text="Browse…", width=120, command=self._on_browse)
        browse_btn.grid(row=0, column=1, sticky="e", padx=(0, 10), pady=10)

        # Sheet name
        sheet_frame = ctk.CTkFrame(self.left)
        sheet_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=8)
        sheet_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(sheet_frame, text="Sheet Name:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.sheet_var = ctk.StringVar(value="Sheet1")
        self.sheet_entry = ctk.CTkEntry(sheet_frame, textvariable=self.sheet_var)
        self.sheet_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        # API configuration
        api_frame = ctk.CTkFrame(self.left)
        api_frame.grid(row=3, column=0, sticky="ew", padx=12, pady=8)
        api_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(api_frame, text="OpenAI API Key:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 6))
        self.api_key_var = ctk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
        self.api_key_entry = ctk.CTkEntry(api_frame, textvariable=self.api_key_var, show="•")
        self.api_key_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=(10, 6))

        ctk.CTkLabel(api_frame, text="Model:").grid(row=1, column=0, sticky="w", padx=10, pady=(6, 10))
        self.model_var = ctk.StringVar(value="gpt-5.2")
        self.model_entry = ctk.CTkEntry(api_frame, textvariable=self.model_var)
        self.model_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=(6, 10))

        # Style strictness
        style_frame = ctk.CTkFrame(self.left)
        style_frame.grid(row=4, column=0, sticky="ew", padx=12, pady=8)
        style_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(style_frame, text="Style Strictness:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.strictness_var = ctk.StringVar(value="Loose")
        self.strictness_menu = ctk.CTkOptionMenu(
            style_frame,
            variable=self.strictness_var,
            values=["Strict", "Loose", "None"],
        )
        self.strictness_menu.grid(row=0, column=1, sticky="w", padx=10, pady=10)

        # Economy (optional)
        econ_frame = ctk.CTkFrame(self.left)
        econ_frame.grid(row=5, column=0, sticky="ew", padx=12, pady=8)
        econ_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(econ_frame, text="Economy (optional):").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.economy_var = ctk.StringVar(value="")
        self.econ_entry = ctk.CTkEntry(econ_frame, textvariable=self.economy_var)
        self.econ_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        # Example text
        example_frame = ctk.CTkFrame(self.left)
        example_frame.grid(row=6, column=0, sticky="nsew", padx=12, pady=8)
        example_frame.grid_columnconfigure(0, weight=1)
        example_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(example_frame, text="Example Text (optional):").grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 6)
        )
        self.example_text = ctk.CTkTextbox(example_frame, height=120, wrap="word")
        self.example_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.example_text.insert("1.0", "")

        # Editing requirements
        edit_frame = ctk.CTkFrame(self.left)
        edit_frame.grid(row=7, column=0, sticky="nsew", padx=12, pady=8)
        edit_frame.grid_columnconfigure(0, weight=1)
        edit_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(edit_frame, text="Editing Requirements (optional):").grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 6)
        )
        self.edit_requirements = ctk.CTkTextbox(edit_frame, height=120, wrap="word")
        self.edit_requirements.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.edit_requirements.insert("1.0", "")

        # Data warning label
        warn = ctk.CTkLabel(
            self.left,
            text="Note: Please ensure your data is properly formatted before processing: time in the first column, headers in the first row, and data in all remaining cells.",
            text_color=("gray35", "gray75"),
            wraplength=420,
            justify="left",
        )
        warn.grid(row=8, column=0, sticky="w", padx=12, pady=(6, 2))

        # Status + controls
        controls = ctk.CTkFrame(self.left)
        controls.grid(row=9, column=0, sticky="ew", padx=12, pady=(8, 12))
        controls.grid_columnconfigure(0, weight=1)

        self.status_var = ctk.StringVar(value="Ready.")
        self.status_label = ctk.CTkLabel(controls, textvariable=self.status_var, anchor="w")
        self.status_label.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))

        self.progress = ctk.CTkProgressBar(controls, mode="indeterminate")
        self.progress.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.progress.stop()

        self.generate_btn = ctk.CTkButton(controls, text="Generate Summary", command=self._on_generate)
        self.generate_btn.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))

    def _build_right_panel(self) -> None:
        self.right = ctk.CTkFrame(self)
        self.right.grid(row=0, column=1, sticky="nsew", padx=(0, 12), pady=12)
        self.right.grid_columnconfigure(0, weight=1)
        self.right.grid_rowconfigure(2, weight=1)

        title = ctk.CTkLabel(self.right, text="Output", font=ctk.CTkFont(size=18, weight="bold"))
        title.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        btn_row = ctk.CTkFrame(self.right, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        btn_row.grid_columnconfigure(0, weight=1)

        self.edit_btn = ctk.CTkButton(btn_row, text="Edit Summary", width=140, command=self._on_edit_summary)
        self.edit_btn.grid(row=0, column=0, sticky="w")
        self.edit_btn.configure(state="disabled")

        self.copy_btn = ctk.CTkButton(btn_row, text="Copy to Clipboard", width=160, command=self._on_copy)
        self.copy_btn.grid(row=0, column=1, sticky="e")
        self.copy_btn.configure(state="disabled")

        self.output_box = ctk.CTkTextbox(self.right, wrap="word")
        self.output_box.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.output_box.insert("1.0", self._placeholder_text + "\n")

    def _has_real_output(self) -> bool:
        text = self.output_box.get("1.0", "end").strip()
        return bool(text) and text != self._placeholder_text

    def _set_busy(self, busy: bool, status: str) -> None:
        self.status_var.set(status)
        if busy:
            self.generate_btn.configure(state="disabled")
            self.copy_btn.configure(state="disabled")
            self.edit_btn.configure(state="disabled")
            self.progress.start()
        else:
            self.generate_btn.configure(state="normal")
            enable = "normal" if self._has_real_output() else "disabled"
            self.copy_btn.configure(state=enable)
            self.edit_btn.configure(state=enable)
            self.progress.stop()

    def _on_browse(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file_path_var.set(path)

    def _validate_api_inputs(self) -> Tuple[str, str]:
        api_key = (self.api_key_var.get() or "").strip()
        model = (self.model_var.get() or "").strip() or "gpt-5.2"
        if not api_key:
            raise ValueError("Please enter your OpenAI API key.")
        return api_key, model

    def _validate_inputs(self) -> Tuple[str, str, str, str, str, str]:
        path = (self.file_path_var.get() or "").strip()
        sheet = (self.sheet_var.get() or "").strip() or "Sheet1"
        api_key = (self.api_key_var.get() or "").strip()
        model = (self.model_var.get() or "").strip() or "gpt-5.2"
        strictness = (self.strictness_var.get() or "None").strip()
        economy = (self.economy_var.get() or "").strip()

        if not path:
            raise ValueError("Please select an Excel file.")
        if not os.path.exists(path):
            raise ValueError("The selected Excel file path does not exist.")
        if not api_key:
            raise ValueError("Please enter your OpenAI API key.")
        if strictness not in ("Strict", "Loose", "None"):
            raise ValueError("Style Strictness must be Strict, Loose, or None.")

        return path, sheet, api_key, model, strictness, economy

    def _on_generate(self) -> None:
        if self._worker_thread and self._worker_thread.is_alive():
            messagebox.showinfo("In progress", "A generation task is already running.")
            return

        try:
            path, sheet, api_key, model, strictness, economy = self._validate_inputs()
        except Exception as e:
            messagebox.showwarning("Missing/Invalid Input", str(e))
            return

        example = self.example_text.get("1.0", "end").strip()

        self._set_busy(True, "Loading Excel and generating summary (Step 1/2, then Step 2/2)…")

        def worker() -> None:
            try:
                excel_slice = load_excel_slice(path, sheet_name=sheet, last_n_rows=24)

                if excel_slice.warning:
                    # Send a non-fatal warning via queue (status channel)
                    self._result_queue.put(("warning", excel_slice.warning, None))

                step1_json, paragraph = run_two_step_pipeline(
                    api_key=api_key,
                    model=model,
                    markdown_table=excel_slice.markdown_table,
                    example_text=example,
                    strictness=strictness,
                    economy_input=economy,
                )

                # We display only the final paragraph by default; keep step1 in case you want it later.
                payload = json.dumps({"paragraph": paragraph, "step1": step1_json}, ensure_ascii=False)
                self._result_queue.put(("success", payload, None))

            except Exception as e:
                details = "".join(traceback.format_exception(type(e), e, e.__traceback__))
                self._result_queue.put(("error", str(e), details))

        self._worker_thread = threading.Thread(target=worker, daemon=True)
        self._worker_thread.start()

    def _on_edit_summary(self) -> None:
        if self._worker_thread and self._worker_thread.is_alive():
            messagebox.showinfo("In progress", "A generation or edit task is already running.")
            return

        current_summary = self.output_box.get("1.0", "end").strip()
        if not current_summary or current_summary == self._placeholder_text:
            messagebox.showwarning("No Summary", "Please generate a summary before editing.")
            return

        try:
            api_key, model = self._validate_api_inputs()
        except Exception as e:
            messagebox.showwarning("Missing/Invalid Input", str(e))
            return

        requirements = self.edit_requirements.get("1.0", "end").strip()

        self._set_busy(True, "Editing summary based on your requirements...")

        def worker() -> None:
            try:
                edited = run_edit_pipeline(
                    api_key=api_key,
                    model=model,
                    original_summary=current_summary,
                    edit_requirements=requirements,
                )
                self._result_queue.put(("edit_success", edited, None))
            except Exception as e:
                details = "".join(traceback.format_exception(type(e), e, e.__traceback__))
                self._result_queue.put(("error", str(e), details))

        self._worker_thread = threading.Thread(target=worker, daemon=True)
        self._worker_thread.start()

    def _on_copy(self) -> None:
        text = self.output_box.get("1.0", "end").strip()
        if not text:
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status_var.set("Copied to clipboard.")

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, msg, details = self._result_queue.get_nowait()

                if kind == "warning":
                    # Non-fatal status update
                    self.status_var.set(f"Warning: {msg}")

                elif kind == "success":
                    try:
                        obj = json.loads(msg)
                        paragraph = obj.get("paragraph", "").strip()
                    except Exception:
                        paragraph = msg.strip()

                    self.output_box.delete("1.0", "end")
                    self.output_box.insert("1.0", paragraph + "\n")
                    self._set_busy(False, "Done.")

                elif kind == "edit_success":
                    paragraph = msg.strip()
                    self.output_box.delete("1.0", "end")
                    self.output_box.insert("1.0", paragraph + "\n")
                    self._set_busy(False, "Edited.")

                elif kind == "error":
                    self._set_busy(False, "Failed.")
                    if details:
                        # Keep messagebox readable; provide detail as optional expansion-ish text
                        messagebox.showerror("Error", f"{msg}\n\nDetails (for debugging):\n{details[:4000]}")
                    else:
                        messagebox.showerror("Error", msg)

        except queue.Empty:
            pass
        finally:
            self.after(120, self._poll_queue)


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
