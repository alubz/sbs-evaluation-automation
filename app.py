import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from functools import lru_cache
from textwrap import wrap

import matplotlib
matplotlib.use("Agg")  # headless backend: no GUI, faster, thread/server safe
from matplotlib.figure import Figure
import numpy as np
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# Constants
CORE_COLUMNS = ['Program Name', 'Faculty Name', 'Session', 'Category', 'Comments']
TEXT_KEY_COLUMNS = ['Program Name', 'Faculty Name', 'Session', 'Department']
DEPT_COL = "Department"  # optional column: which department each programme belongs to
UNSPECIFIED_SESSION = "Unspecified session"  # label for rows that have no Session value
SIGNOFF_FILE = "signoff.xlsx"  # Excel file with Name | Position | Sign, kept next to the app
SIGNATURE_MAX_HEIGHT_IN = 0.6  # default signature picture box in inches: the picture is scaled to fit
SIGNATURE_MAX_WIDTH_IN = 2.0   # inside it, keeping its proportions
SIGNATURE_LIMIT_IN = 6.0       # absolute cap so a typo in Excel cannot blow the picture up
EXCEL_SHEET = "Comments & Scores"
LOGO_CANDIDATES = ("logo.png", "input_file_0.png", "SBS_Logo.png")
SCORE_MAX = 5
EXCEL_BLUE = "#4472C4"
EXCEL_RED = "#C00000"  # Professional Excel dark red for charts
CHART_DPI = 150

# Master Departmental Report settings
RANK_N = 3                       # default number of top / bottom faculty sessions shown
MAX_COMMENTS_PER_FACULTY = 5     # default cap on comments per session (None = all)
MIN_RESPONSES_TO_RANK = 3        # sessions with fewer responses are ignored when ranking
MAX_COMMENT_CHARS = 300          # long comments are trimmed to this length


def get_evaluation_params(df):
    """Dynamically detect evaluation parameter columns from the dataframe, excluding metadata and year columns."""
    return [
        col for col in df.columns
        if col not in CORE_COLUMNS
        and col != DEPT_COL
        and 'year' not in col.lower()
        and pd.api.types.is_numeric_dtype(df[col])
    ]

def validate_dataframe(df):
    """Validate that required core columns exist in the dataframe."""
    missing_cols = [col for col in CORE_COLUMNS if col not in df.columns]
    if missing_cols:
        return False, f"Missing required columns: {', '.join(missing_cols)}"
    return True, None

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise headers and key text columns so groups don't split on stray spaces."""
    df = df.dropna(how="all").copy()
    df.columns = df.columns.astype(str).str.strip()
    for col in TEXT_KEY_COLUMNS:
        if col in df.columns:
            s = df[col]
            df[col] = s.where(s.isna(), s.astype(str).str.strip()).replace("", np.nan)
    return df

def _has_dept(df) -> bool:
    return DEPT_COL in df.columns and df[DEPT_COL].notna().any()

@st.cache_data(show_spinner="Reading file…")
def load_uploaded_file(file_bytes: bytes, name: str) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    if name.lower().endswith(".csv"):
        return clean_dataframe(pd.read_csv(buf))
    xls = pd.ExcelFile(buf)
    if EXCEL_SHEET not in xls.sheet_names:
        raise ValueError(f"Sheet '{EXCEL_SHEET}' not found. Sheets in file: {', '.join(xls.sheet_names)}")
    return clean_dataframe(xls.parse(EXCEL_SHEET, header=1))

def _read_google_csv(export_url):
    """Read the sheet; if the headers are on the second row (like the Excel layout), retry with header=1."""
    df = clean_dataframe(pd.read_csv(export_url))
    if not validate_dataframe(df)[0]:
        try:
            alt = clean_dataframe(pd.read_csv(export_url, header=1))
            if validate_dataframe(alt)[0]:
                return alt
        except Exception:
            pass
    return df

@st.cache_data(show_spinner="Fetching Google Sheet…", ttl=300)
def load_google_sheet(sheet_url: str):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        return None, "Invalid Google Sheets URL format."
    gid_match = re.search(r"gid=([0-9]+)", sheet_url)
    gid = gid_match.group(1) if gid_match else "0"
    export_url = f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv&gid={gid}"
    try:
        return _read_google_csv(export_url), None
    except Exception as e:
        return None, str(e)

# Dataclasses for structured data
@dataclass
class FacultyResult:
    session: str
    faculty: str
    respondents: int
    params: list
    scores: list
    prog_avgs: list
    avg_score: float
    avg_prog: float
    comments: list
    chart: bytes = b""

@dataclass
class DeptResult:
    session: str
    params: list
    scores: list
    avg_score: float
    comments: list
    chart: bytes = b""
    respondents: int = 0

@dataclass
class ProgrammeResult:
    faculty: list = field(default_factory=list)
    dept: list = field(default_factory=list)
    highest_session_respondents: int = 0

@dataclass
class Signatory:
    name: str = ""
    title: str = ""
    signature: bytes = b""   # image of the signature (picture placed in the Sign cell)
    sign_text: str = ""      # or a typed signature
    sign_width: float = 0.0  # optional picture width in inches (0 = automatic)
    sign_height: float = 0.0 # optional picture height in inches (0 = automatic)

def _num(value) -> float:
    try:
        number = float(str(value).replace(",", ".").strip())
        if number <= 0:
            return 0.0
        return min(number, SIGNATURE_LIMIT_IN)
    except (TypeError, ValueError):
        return 0.0

def _read_in_cell_signatures(workbook_bytes, worksheet_name, signature_col):
    """Read Excel's newer in-cell images, which openpyxl does not expose as worksheet images."""
    from posixpath import join, normpath
    from xml.etree import ElementTree as ET
    from zipfile import ZipFile

    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rich_value_rel_ns = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
    rich_data_ns = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"

    with ZipFile(io.BytesIO(workbook_bytes)) as archive:
        names = set(archive.namelist())
        required = {
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/metadata.xml",
            "xl/richData/rdrichvalue.xml",
            "xl/richData/richValueRel.xml",
            "xl/richData/_rels/richValueRel.xml.rels",
        }
        if not required.issubset(names):
            return {}

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        workbook_rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.get("Id"): rel.get("Target")
            for rel in workbook_rels.findall(f"{{{package_rel_ns}}}Relationship")
        }
        sheet = next(
            (node for node in workbook.findall(f".//{{{main_ns}}}sheet")
             if node.get("name") == worksheet_name),
            None,
        )
        if sheet is None:
            return {}
        sheet_target = rel_targets.get(sheet.get(f"{{{rel_ns}}}id"))
        if not sheet_target:
            return {}
        sheet_path = normpath(sheet_target.lstrip("/")) if sheet_target.startswith("/") else normpath(join("xl", sheet_target))
        if sheet_path not in names:
            return {}

        metadata = ET.fromstring(archive.read("xl/metadata.xml"))
        metadata_blocks = metadata.find(f"{{{main_ns}}}valueMetadata")
        rich_values = ET.fromstring(archive.read("xl/richData/rdrichvalue.xml"))
        image_rels = ET.fromstring(archive.read("xl/richData/richValueRel.xml"))
        rich_rels = ET.fromstring(archive.read("xl/richData/_rels/richValueRel.xml.rels"))
        if metadata_blocks is None:
            return {}

        metadata_records = list(metadata_blocks.findall(f"{{{main_ns}}}bk"))
        value_records = list(rich_values)
        image_rel_ids = [
            item.get(f"{{{rel_ns}}}id")
            for item in image_rels.findall(f"{{{rich_value_rel_ns}}}rel")
        ]
        rich_rel_targets = {
            rel.get("Id"): rel.get("Target")
            for rel in rich_rels.findall(f"{{{package_rel_ns}}}Relationship")
        }
        worksheet = ET.fromstring(archive.read(sheet_path))
        signatures = {}
        for cell in worksheet.findall(f".//{{{main_ns}}}c"):
            reference = cell.get("r", "")
            cell_match = re.fullmatch(r"([A-Z]+)(\d+)", reference)
            if not cell_match or cell_match.group(1) != signature_col:
                continue
            metadata_index = cell.get("vm")
            if not metadata_index:
                continue
            metadata_index = int(metadata_index) - 1
            if not 0 <= metadata_index < len(metadata_records):
                continue
            rich_ref = metadata_records[metadata_index].find(f".//{{{main_ns}}}rc")
            if rich_ref is None or rich_ref.get("v") is None:
                continue
            value_index = int(rich_ref.get("v"))
            if not 0 <= value_index < len(value_records):
                continue
            image_id = value_records[value_index].find(f"{{{rich_data_ns}}}v")
            if image_id is None or image_id.text is None:
                continue
            image_index = int(image_id.text)
            if not 0 <= image_index < len(image_rel_ids):
                continue
            target = rich_rel_targets.get(image_rel_ids[image_index])
            if not target:
                continue
            image_path = (normpath(target.lstrip("/")) if target.startswith("/")
                          else normpath(join("xl/richData", target)))
            if image_path in names:
                signatures[int(cell_match.group(2))] = archive.read(image_path)
        return signatures

def read_signoffs(source) -> list:
    """Read the sign-off Excel. Headers in row 1: Name | Position | Sign (one row per person).
    The Sign cell can hold typed text or a picture placed in the cell or on that row."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    try:
        if isinstance(source, bytes):
            workbook_bytes = source
        elif hasattr(source, "getvalue"):
            workbook_bytes = source.getvalue()
        else:
            with open(source, "rb") as workbook_file:
                workbook_bytes = workbook_file.read()
        wb = load_workbook(io.BytesIO(workbook_bytes))
    except Exception as exc:
        raise ValueError("The sign-off file could not be opened as a valid Excel workbook.") from exc
    ws = next((wb[n] for n in wb.sheetnames if n.strip().lower().replace(" ", "").replace("-", "") == "signoff"),
              wb.active)

    cols = {}
    for c in ws[1]:
        h = re.sub(r"\(.*?\)", "", str(c.value or "").strip().lower()).strip()   # ignore "(in)" etc.
        if h == "name":
            cols["name"] = c.column
        elif h in ("position", "title", "role", "designation"):
            cols["title"] = c.column
        elif h in ("sign", "signature"):
            cols["sign"] = c.column
        elif h in ("sign width", "signature width", "width"):
            cols["w"] = c.column
        elif h in ("sign height", "signature height", "height"):
            cols["h"] = c.column
    if "name" not in cols:
        raise ValueError("Row 1 needs the headers: Name, Position, Sign")
    # Floating pictures, matched to the row they are placed on.
    pics = {}
    for img in getattr(ws, "_images", []):
        try:
            anchor = img.anchor
            row = (anchor._from.row + 1) if hasattr(anchor, "_from") else None
            data = img._data()
            if row is not None and data:
                pics.setdefault(row, data)
        except Exception:
            continue
    if "sign" in cols:
        pics.update(_read_in_cell_signatures(
            workbook_bytes, ws.title, get_column_letter(cols["sign"])))

    people = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, cols["name"]).value or "").strip()
        title = str(ws.cell(r, cols["title"]).value or "").strip() if "title" in cols else ""
        sign_cell = ws.cell(r, cols["sign"]) if "sign" in cols else None
        text = str(sign_cell.value or "").strip() if sign_cell is not None else ""
        if sign_cell is not None and sign_cell.data_type == "e":
            text = ""
        if not (name or title or text or r in pics):
            continue
        people.append(Signatory(
            name, title, pics.get(r, b""), text,
            _num(ws.cell(r, cols["w"]).value) if "w" in cols else 0.0,
            _num(ws.cell(r, cols["h"]).value) if "h" in cols else 0.0))
    return people

def signoff_template_bytes() -> bytes:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Sign-off"
    ws.append(["Name", "Position", "Sign", "Sign Width (in)", "Sign Height (in)"])
    ws.append(["Jane Doe", "Head of Quality Assurance", "Jane D.", None, None])
    ws["G1"] = "Notes"
    ws["G2"] = "Sign Width / Height are optional (inches) and only apply to a signature picture."
    ws["G3"] = "Fill one to keep the proportions, or both to set the exact size. Leave empty for the default size."
    ws.column_dimensions["G"].width = 90
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# Helper functions
def _to_pct(x):
    return (x / SCORE_MAX * 100).round(0)

def _pct_str(x):
    return f"{int(round(x))}%" if pd.notna(x) else "N/A"

def _fmt_count(n):
    """Headcounts that were not entered show as 'Not provided'."""
    return str(int(n)) if n is not None and pd.notna(n) else "Not provided"

def _max_session_responses(prog_df: pd.DataFrame) -> int:
    """Largest number of responses in any one session, counting faculty AND department sessions."""
    sizes = []
    for part in (prog_df[prog_df["Faculty Name"].notna()], prog_df[prog_df["Faculty Name"].isna()]):
        if not part.empty:
            sizes.append(int(part.groupby("Session").size().max()))
    return max(sizes) if sizes else 0

# Analysis functions
def analyse_programme(prog_df: pd.DataFrame, param_cols: list) -> ProgrammeResult:
    result = ProgrammeResult()
    result.highest_session_respondents = _max_session_responses(prog_df)
    fac_rows = prog_df[prog_df["Faculty Name"].notna()]
    other_rows = prog_df[prog_df["Faculty Name"].isna()]

    if not fac_rows.empty:
        grouped = fac_rows.groupby(["Session", "Faculty Name"], sort=False)   # sheet order
        scores = _to_pct(grouped[param_cols].mean())
        prog_avg = scores.mean().round(0)
        score_rows = scores.to_dict("index")

        for (session, faculty), group in grouped:
            row = score_rows[(session, faculty)]
            params = [p for p in param_cols if pd.notna(row[p])]
            ind = [row[p] for p in params]
            pav = [prog_avg[p] for p in params]
            result.faculty.append(FacultyResult(
                session=str(session), faculty=str(faculty), respondents=len(group),
                params=params, scores=ind, prog_avgs=pav,
                avg_score=round(sum(ind) / len(ind), 0) if ind else 0,
                avg_prog=round(sum(pav) / len(pav), 0) if pav else 0,
                comments=group["Comments"].dropna().tolist(),
            ))

    for session, group in other_rows.groupby("Session", sort=False):   # sheet order
        params = [p for p in param_cols if group[p].notna().any()]
        ind = _to_pct(group[params].mean()).tolist() if params else []
        result.dept.append(DeptResult(
            session=str(session), params=params, scores=ind,
            avg_score=round(sum(ind) / len(ind), 0) if ind else 0,
            comments=group["Comments"].dropna().tolist(),
            respondents=len(group),
        ))
    return result

# Chart rendering (memory-based)
def render_chart(params, series, title, series_labels, colors=None) -> bytes:
    """Excel-style clustered column chart. Returns PNG bytes."""
    colors = colors or [EXCEL_BLUE, EXCEL_RED]
    n = len(series)

    fig = Figure(figsize=(12, 6.5), dpi=CHART_DPI)
    fig.patch.set_facecolor("white")
    ax = fig.subplots()

    x = np.arange(len(params))
    bar_w = 0.5 / n

    valid_series = [np.asarray(s, dtype=float) for s in series if len(s) > 0]
    if valid_series:
        all_vals = np.concatenate(valid_series)
        min_val = int(np.nanmin(all_vals)) if not np.isnan(all_vals).all() else 0
        ymin = max(0, min_val - 4)
    else:
        ymin = 0

    for i, scores in enumerate(series):
        offset = (i - (n - 1) / 2) * bar_w
        bars = ax.bar(x + offset, scores, bar_w, color=colors[i % len(colors)], label=series_labels[i], edgecolor='white', linewidth=1)

        formatted_labels = [
            f"{int(round(v))}%" if (pd.notna(v) and not np.isnan(v)) else "0%"
            for v in scores
        ]
        ax.bar_label(bars, labels=formatted_labels, padding=4, fontsize=9, fontweight='bold')

    ax.set_ylim(0, 135)
    ax.set_yticks(range(0, 140, 10))
    ax.tick_params(axis="y", labelsize=11)

    wrap_width = max(12, min(22, int(140 / max(len(params), 1))))
    wrapped_params = ["\n".join(wrap(str(p), wrap_width)) for p in params]

    ax.set_xticks(x)
    ax.set_xticklabels(wrapped_params, fontsize=10, rotation=45, ha='right', rotation_mode="anchor")
    ax.tick_params(axis="x", length=0)

    ax.set_ylabel("% Score", fontsize=12, fontweight="bold")
    ax.set_xlabel("Evaluation Parameters", fontsize=12, fontweight="bold", labelpad=10)
    ax.set_title(title, fontsize=13, fontweight="bold", pad=15)

    ax.grid(axis='y', alpha=0.3, linestyle='--')
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_linewidth(0.5)

    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles, labels, loc="upper right", fontsize=11, framealpha=0.9)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=CHART_DPI, bbox_inches="tight", pad_inches=0.4)
    return buf.getvalue()

# Word document helpers
@lru_cache(maxsize=1)
def _logo_bytes():
    for fname in LOGO_CANDIDATES:
        if os.path.exists(fname):
            with open(fname, "rb") as f:
                return f.read()
    return None

def new_doc() -> Document:
    doc = Document()
    logo = _logo_bytes()
    if logo:
        hp = doc.sections[0].header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        try:
            hp.add_run().add_picture(io.BytesIO(logo), width=Inches(1.5))
        except Exception:
            pass
    return doc

def doc_to_bytes(doc: Document) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def safe_name(text, max_len=40) -> str:
    return re.sub(r"[\\/:*?\"<>|'()\r\n]+", "_", str(text)).strip()[:max_len] or "unnamed"

def add_title(doc, text, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)

def _fill(cell, text, bold=False, size=9, center=False):
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Enhanced Professional Corporate Table Styling
def set_cell_background(cell, fill_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=100, bottom=100, start=150, end=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', start), ('right', end)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_score_table(doc, params, scores, avg, prog_avgs=None, avg_prog=None, avg_label="Average scores"):
    compare = prog_avgs is not None
    headers = ["Evaluation Parameter", "(%) Individual Score" if compare else "(%) Score"]
    if compare:
        headers.append("(%) Programme Average")

    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header Row Styling
    hdr_cells = table.rows[0].cells
    for i, cell in enumerate(hdr_cells):
        set_cell_background(cell, "1F4E78")  # Dark Executive Blue
        set_cell_margins(cell, top=140, bottom=140, start=180, end=180)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(headers[i])
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)

    # Data Rows with Zebra Striping
    for idx, name in enumerate(params):
        row_cells = table.add_row().cells
        bg_color = "F9FAFC" if idx % 2 == 1 else "FFFFFF"

        for cell in row_cells:
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=100, bottom=100, start=180, end=180)

        _fill(row_cells[0], name, size=9, center=False)
        s_val = f"{int(round(scores[idx]))}%" if pd.notna(scores[idx]) else "N/A"
        _fill(row_cells[1], s_val, center=True, size=9)

        if compare:
            p_val = f"{int(round(prog_avgs[idx]))}%" if pd.notna(prog_avgs[idx]) else "N/A"
            _fill(row_cells[2], p_val, center=True, size=9)

    # Summary / Footer Row
    total_cells = table.add_row().cells
    for cell in total_cells:
        set_cell_background(cell, "EBF1F5")
        set_cell_margins(cell, top=120, bottom=120, start=180, end=180)

    _fill(total_cells[0], avg_label, bold=True, size=9.5)
    avg_val = f"{int(round(avg))}%" if pd.notna(avg) else "0%"
    _fill(total_cells[1], avg_val, bold=True, center=True, size=9.5)

    if compare:
        avg_prog_val = f"{int(round(avg_prog))}%" if pd.notna(avg_prog) else "0%"
        _fill(total_cells[2], avg_prog_val, bold=True, center=True, size=9.5)

def add_styled_table(doc, headers, rows):
    """Generic table with the same navy header and zebra striping as add_score_table."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, cell in enumerate(table.rows[0].cells):
        set_cell_background(cell, "1F4E78")
        set_cell_margins(cell, top=140, bottom=140, start=180, end=180)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(headers[i])
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)
    for idx, row in enumerate(rows):
        cells = table.add_row().cells
        bg = "F9FAFC" if idx % 2 else "FFFFFF"
        for j, (cell, val) in enumerate(zip(cells, row)):
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=100, bottom=100, start=180, end=180)
            _fill(cell, val, size=9, center=j > 0)

def add_comments(doc, comments, level):
    doc.add_heading("Comments & Feedback", level=level)
    if comments:
        for c in comments:
            doc.add_paragraph(str(c), style="List Bullet")
    else:
        doc.add_paragraph("No comments recorded.")

def _signature_size(img_bytes, width_in=0.0, height_in=0.0):
    """Width and height (inches) for the pasted signature.
    Only one of width/height given: the other follows the picture's proportions.
    Both given: used as is. Neither: scaled to fit inside the default box, proportions kept."""
    try:
        from PIL import Image as PILImage
        with PILImage.open(io.BytesIO(img_bytes)) as im:
            ratio = im.size[0] / im.size[1]
    except Exception:
        ratio = None
    if ratio is None:                                   # cannot read the picture: fixed height, automatic width
        return None, (height_in or SIGNATURE_MAX_HEIGHT_IN)
    if width_in and height_in:
        w, h = width_in, height_in
    elif height_in:
        w, h = height_in * ratio, height_in
    elif width_in:
        w, h = width_in, width_in / ratio
    else:
        h = SIGNATURE_MAX_HEIGHT_IN
        w = h * ratio
        if w > SIGNATURE_MAX_WIDTH_IN:
            w = SIGNATURE_MAX_WIDTH_IN
            h = w / ratio
    scale = min(1.0, SIGNATURE_LIMIT_IN / max(w, h))     # never larger than the cap
    return w * scale, h * scale

def add_signature(doc, signatory=None):
    """Sign-off block. Name / Title / Sign are filled from the sign-off details; blank lines if none were given."""
    s = signatory or Signatory()
    blank = "__________________________"
    doc.add_paragraph()
    doc.add_heading("SUBMITTED THROUGH:", level=2)

    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = (Inches(1.25), Inches(5.25))
    fields = (("Name", s.name or blank), ("Title", s.title or blank), ("Sign", None))
    for row, (label, value) in zip(table.rows, fields):
        label_cell, value_cell = row.cells
        for cell, width in zip(row.cells, widths):
            cell.width = width
            set_cell_margins(cell, top=100, bottom=100, start=120, end=120)
        label_run = label_cell.paragraphs[0].add_run(label)
        label_run.bold = True
        if value is not None:
            value_cell.paragraphs[0].add_run(value)

    p = table.rows[2].cells[1].paragraphs[0]
    placed = False
    if s.signature:
        try:
            w, h = _signature_size(s.signature, s.sign_width, s.sign_height)
            p.add_run().add_picture(io.BytesIO(s.signature),
                                    width=Inches(w) if w else None, height=Inches(h))
            placed = True
        except Exception:
            placed = False
    if not placed and s.sign_text:
        run = p.add_run(s.sign_text)
        run.italic = True
        run.font.size = Pt(14)
        placed = True
    if not placed:
        p.add_run(blank)
    doc.add_paragraph(f"Date: {datetime.now().strftime('%B %d, %Y')}")

def add_faculty_content(doc, fr: FacultyResult, pic_width, comment_level):
    if fr.params:
        add_score_table(doc, fr.params, fr.scores, fr.avg_score, fr.prog_avgs, fr.avg_prog)
        if len(fr.params) > 1 and fr.chart:   # a one-bar chart adds nothing the table doesn't say
            doc.add_paragraph()
            doc.add_picture(io.BytesIO(fr.chart), width=Inches(pic_width))
    elif not fr.comments:
        doc.add_paragraph("No evaluation data available for this session.")
    add_comments(doc, fr.comments, comment_level)

# Report builders
def build_faculty_report(programme, faculty_results, session_participants, signatory=None,
                         include_signoff=True) -> bytes:
    doc = new_doc()
    add_title(doc, "FACULTY EVALUATION")
    doc.add_paragraph(f"Programme Name: {programme}")
    doc.add_paragraph(f"Faculty Name: {faculty_results[0].faculty}")
    for index, fr in enumerate(faculty_results):
        if index:
            doc.add_page_break()
        doc.add_heading(f"Session: {fr.session}", level=1)
        doc.add_paragraph(f"No. of Respondents: {fr.respondents}")
        doc.add_paragraph(f"No. of Participants: {_fmt_count(session_participants.get(fr.session))}")
        add_faculty_content(doc, fr, 6.5, 2)
    if include_signoff:
        add_signature(doc, signatory)
    return doc_to_bytes(doc)

def build_overall_report(programme, res: ProgrammeResult, total_participants, dept_headcounts=None,
                         signatory=None, include_signoff=True) -> bytes:
    doc = new_doc()
    highest = res.highest_session_respondents
    if total_participants:
        pct_resp = f"{min(100, round(highest / total_participants * 100))}%"   # never above 100%
    else:
        pct_resp = "Not available (participants not provided)"

    add_title(doc, f"OVERALL PROGRAMME EVALUATION REPORT\n{programme}")
    doc.add_paragraph(f"No. of Respondents: {highest}")
    doc.add_paragraph(f"No. of Participants: {_fmt_count(total_participants)}")
    doc.add_paragraph(f"Response Rate: {pct_resp}")
    doc.add_paragraph("-" * 50)

    if res.faculty:
        doc.add_heading("Faculty Reports", level=1)
        for fr in res.faculty:
            doc.add_heading(f"{fr.faculty} - {fr.session}", level=2)
            add_faculty_content(doc, fr, 6.5, 3)
            doc.add_page_break()

    if res.dept:
        doc.add_heading("Department Reports", level=1)
        for d in res.dept:
            doc.add_heading(d.session, level=2)
            line = f"No. of Respondents: {d.respondents}"
            headcount = (dept_headcounts or {}).get(d.session)
            if headcount:
                rate = min(100, round(d.respondents / headcount * 100))
                line += f"  |  No. of Participants: {int(headcount)}  |  Response Rate: {rate}%"
            doc.add_paragraph(line)
            if d.params:
                add_score_table(doc, d.params, d.scores, d.avg_score, avg_label="Average score")
                if len(d.params) > 1 and d.chart:   # skip the chart when there is only one parameter
                    doc.add_paragraph()
                    doc.add_picture(io.BytesIO(d.chart), width=Inches(6.5))
            add_comments(doc, d.comments, 3)
            doc.add_page_break()

    if include_signoff:
        add_signature(doc, signatory)
    return doc_to_bytes(doc)

# Master Departmental Report (high-level summary for directors and heads)
def _rank_faculty(fac, param_cols, rank_n=RANK_N):
    """Rank each (Programme, Session, Faculty) on overall % score. Returns (top, bottom)."""
    keys = ["Program Name", "Session", "Faculty Name"]
    g = fac.groupby(keys)
    stats = pd.DataFrame({
        "score": g[param_cols].mean().mean(axis=1) / SCORE_MAX * 100,
        "responses": g.size(),
    }).dropna(subset=["score"])

    eligible = stats[stats["responses"] >= MIN_RESPONSES_TO_RANK]
    if len(eligible) < 2 * rank_n:      # too few sessions to apply the threshold
        eligible = stats
    eligible = eligible.sort_values("score", ascending=False)

    top = eligible.head(rank_n)
    bottom = eligible.drop(top.index).tail(rank_n).sort_values("score")  # worst first
    return top, bottom

def _add_faculty_block(doc, key, row, fac, max_comments=MAX_COMMENTS_PER_FACULTY):
    prog, session, name = key
    doc.add_heading(f"{name}: {_pct_str(row['score'])}", level=3)
    doc.add_paragraph(f"{prog}  |  {session}  |  {int(row['responses'])} responses")

    mask = ((fac["Program Name"] == prog) & (fac["Session"] == session)
            & (fac["Faculty Name"] == name))
    comments = fac.loc[mask, "Comments"].dropna().astype(str).tolist()
    if not comments:
        doc.add_paragraph("No comments recorded.")
        return
    shown = comments if max_comments is None else comments[:max_comments]
    for c in shown:
        c = c if len(c) <= MAX_COMMENT_CHARS else c[:MAX_COMMENT_CHARS].rstrip() + "…"
        doc.add_paragraph(c, style="List Bullet")
    if len(comments) > len(shown):
        doc.add_paragraph(f"+ {len(comments) - len(shown)} more comments not shown.")

def build_master_report(df, param_cols, rank_n=RANK_N, max_comments=MAX_COMMENTS_PER_FACULTY,
                        include_support=True, include_chart=True, scope_label=None, signatory=None) -> bytes:
    fac = df[df["Faculty Name"].notna()]
    support = df[df["Faculty Name"].isna()] if include_support else df.iloc[0:0]
    doc = new_doc()

    title = "STRATHMORE BUSINESS SCHOOL\nMASTER DEPARTMENTAL EVALUATION REPORT"
    if scope_label:
        title += f"\n{scope_label}"
    add_title(doc, title)

    if fac.empty:
        doc.add_paragraph("No faculty evaluation records found.")
        return doc_to_bytes(doc)

    has_dept = _has_dept(df)
    n_depts = df[DEPT_COL].nunique() if has_dept else 0
    prog_dept = (df.dropna(subset=[DEPT_COL]).groupby("Program Name")[DEPT_COL].first()
                 if has_dept else pd.Series(dtype=object))

    glob = _to_pct(fac[param_cols].mean()).dropna()
    prog_avg = (fac.groupby("Program Name")[param_cols].mean().mean(axis=1)
                / SCORE_MAX * 100).dropna().sort_values(ascending=False)
    n_resp = fac.groupby("Program Name").size()
    support_avg = (support.groupby("Program Name")[param_cols].mean().mean(axis=1) / SCORE_MAX * 100
                   if not support.empty else pd.Series(dtype=float))

    dept_avg = pd.Series(dtype=float)
    if n_depts > 1:
        dept_avg = (fac.groupby(DEPT_COL)[param_cols].mean().mean(axis=1)
                    / SCORE_MAX * 100).dropna().sort_values(ascending=False)

    top, bottom = _rank_faculty(fac, param_cols, rank_n)
    mixed = sorted(set(top.index.get_level_values("Faculty Name"))
                   & set(bottom.index.get_level_values("Faculty Name")))

    # 1. Snapshot
    doc.add_heading("Snapshot", level=1)
    snap = [datetime.now().strftime('%B %d, %Y')]
    if n_depts:
        snap.append(f"{n_depts} department{'s' if n_depts != 1 else ''}")
    snap += [f"{df['Program Name'].nunique()} programmes",
             f"{len(fac)} faculty responses",
             f"Overall faculty score: {_pct_str(glob.mean())}"]
    doc.add_paragraph("  |  ".join(snap))

    # 2. Highlights
    doc.add_heading("Highlights", level=1)
    if len(dept_avg) > 1:
        doc.add_paragraph(f"Top department: {dept_avg.index[0]} ({_pct_str(dept_avg.iloc[0])})",
                          style="List Bullet")
        doc.add_paragraph(f"Needs most attention (department): {dept_avg.index[-1]} ({_pct_str(dept_avg.iloc[-1])})",
                          style="List Bullet")
    if not prog_avg.empty:
        doc.add_paragraph(f"Top programme: {prog_avg.index[0]} ({_pct_str(prog_avg.iloc[0])})",
                          style="List Bullet")
        if len(prog_avg) > 1:
            doc.add_paragraph(f"Needs most attention (programme): {prog_avg.index[-1]} ({_pct_str(prog_avg.iloc[-1])})",
                              style="List Bullet")
    if not glob.empty:
        ranked = glob.sort_values(ascending=False)
        doc.add_paragraph(f"Strongest parameter: {ranked.index[0]} ({_pct_str(ranked.iloc[0])})",
                          style="List Bullet")
        doc.add_paragraph(f"Weakest parameter: {ranked.index[-1]} ({_pct_str(ranked.iloc[-1])})",
                          style="List Bullet")
    if pd.notna(support_avg.mean()):
        doc.add_paragraph(f"Support services average: {_pct_str(support_avg.mean())}", style="List Bullet")
    if mixed:
        doc.add_paragraph(
            f"Mixed results: {', '.join(mixed)} appear in both the top and bottom lists "
            f"(strong in one session, weak in another).", style="List Bullet")

    # 3. Department scores (only when more than one department is in the report)
    if len(dept_avg) > 1:
        doc.add_heading("Department Scores", level=1)
        n_prog_by_dept = fac.groupby(DEPT_COL)["Program Name"].nunique()
        n_resp_by_dept = fac.groupby(DEPT_COL).size()
        rows = [[d, _pct_str(s), int(n_prog_by_dept.get(d, 0)), int(n_resp_by_dept.get(d, 0))]
                for d, s in dept_avg.items()]
        add_styled_table(doc, ["Department", "Faculty Score", "Programmes", "Responses"], rows)

    # 4. Programme league table
    doc.add_heading("Programme Scores", level=1)
    headers = ["Programme"] + (["Department"] if n_depts > 1 else []) + ["Faculty Score", "Responses"]
    if include_support:
        headers.append("Support Score")
    rows = []
    for p, s in prog_avg.items():
        r = [p]
        if n_depts > 1:
            r.append(prog_dept.get(p, ""))
        r += [_pct_str(s), int(n_resp.get(p, 0))]
        if include_support:
            r.append(_pct_str(support_avg.get(p, np.nan)))
        rows.append(r)
    add_styled_table(doc, headers, rows)

    # 5. One chart (skipped when there is only one parameter)
    if include_chart and len(glob) > 1:
        doc.add_paragraph()
        png = render_chart(glob.index.tolist(), [glob.tolist()],
                           "Faculty Evaluation: Average by Parameter",
                           ["Average"], [EXCEL_BLUE])
        doc.add_picture(io.BytesIO(png), width=Inches(6.0))

    # 6. Top / bottom sessions with comments
    doc.add_heading(f"Top {len(top)} Faculty Sessions", level=1)
    for key, row in top.iterrows():
        _add_faculty_block(doc, key, row, fac, max_comments)
    if not bottom.empty:
        doc.add_heading(f"Bottom {len(bottom)} Faculty Sessions", level=1)
        for key, row in bottom.iterrows():
            _add_faculty_block(doc, key, row, fac, max_comments)

    if signatory:
        add_signature(doc, signatory)
    return doc_to_bytes(doc)

# Batch generation (in-memory, guaranteed valid ZIP structure)
def generate_programme_zip(df, programmes, participants, totals, gen_overall, gen_individual,
                           param_cols, on_progress=None, signatory=None, include_signoff=True,
                           combine_faculty_sessions=True) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used_folders = set()
        for idx, prog in enumerate(programmes):
            prog_df = df[df["Program Name"] == prog]
            res = analyse_programme(prog_df, param_cols)
            folder = _unique(safe_name(prog, 80), used_folders)

            # Render charts once, share between reports (no chart when a session has only one parameter)
            for fr in res.faculty:
                if len(fr.params) > 1:
                    fr.chart = render_chart(fr.params, [fr.scores, fr.prog_avgs],
                                            f"{fr.session} : {fr.faculty}", ["Individual", "Programme Avg"])
            if gen_overall:
                for d in res.dept:
                    if len(d.params) > 1:
                        d.chart = render_chart(d.params, [d.scores], f"{d.session} : Department Evaluation",
                                               ["Department Score"], [EXCEL_BLUE])

            if gen_individual:
                used_files = set()
                faculty_sessions = {}
                for fr in res.faculty:
                    faculty_sessions.setdefault(fr.faculty, []).append(fr)
                report_groups = (faculty_sessions.items() if combine_faculty_sessions else
                                 ((fr.faculty, [fr]) for fr in res.faculty))
                for faculty, session_results in report_groups:
                    session_headcounts = {
                        fr.session: participants.get(f"{prog} -- {faculty} -- {fr.session}")
                        for fr in session_results
                    }
                    data = build_faculty_report(
                        prog, session_results, session_headcounts, signatory, include_signoff)
                    suffix = "" if combine_faculty_sessions else f"_{safe_name(session_results[0].session)}"
                    fname = _unique(f"{safe_name(faculty)}{suffix}.docx", used_files)
                    zf.writestr(f"{folder}/Faculty_Reports/{fname}", data)

            if gen_overall:
                total_headcount = totals.get(prog)
                dept_headcounts = {d.session: participants.get(f"{prog} -- Dept -- {d.session}")
                                   for d in res.dept}
                data = build_overall_report(prog, res, total_headcount, dept_headcounts, signatory, include_signoff)
                zf.writestr(f"{folder}/{folder}_Overall_Programme_Report.docx", data)

            if on_progress:
                on_progress((idx + 1) / len(programmes), prog)

    # Ensure stream pointer is reset to start
    buf.seek(0)
    return buf.getvalue()

def _unique(name, used: set) -> str:
    base, ext = os.path.splitext(name)
    candidate, i = name, 2
    while candidate in used:
        candidate = f"{base}_{i}{ext}"
        i += 1
    used.add(candidate)
    return candidate

# UI
def show_preview(df):
    with st.expander("📋 Preview Data"):
        st.dataframe(df.head(10))
        st.caption(f"Showing first 10 of {len(df)} rows")

def ingest():
    source = st.radio("Select Data Source:", ["Upload Excel / CSV File", "Google Sheets Public Link"])
    df, ok_msg = None, ""

    if source == "Upload Excel / CSV File":
        up = st.file_uploader("Upload Evaluation Dataset", type=["xlsx", "csv"])
        if up is None:
            return None
        try:
            df = load_uploaded_file(up.getvalue(), up.name)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return None
        ok_msg = f"Successfully loaded file with {len(df)} total rows."
    else:
        url = st.text_input("Paste Public Google Sheets Link (Anyone with link can view):")
        if not url:
            return None
        df, err = load_google_sheet(url)
        if err:
            st.error(err)
            return None
        ok_msg = f"Successfully connected to Google Sheet ({len(df)} rows loaded)."

    valid, msg = validate_dataframe(df)
    if not valid:
        st.error(msg)
        st.warning("Available columns: " + ", ".join(df.columns.tolist()))
        return None
    st.success(ok_msg)

    # Data-quality checks so responses never disappear silently
    n_no_prog = int(df["Program Name"].isna().sum())
    if n_no_prog:
        st.warning(f"{n_no_prog} rows have no Program Name and will be ignored.")
    n_no_session = int(df.loc[df["Program Name"].notna(), "Session"].isna().sum())
    if n_no_session:
        st.warning(f"{n_no_session} rows have no Session. They are kept under "
                   f"'{UNSPECIFIED_SESSION}' so their responses and comments are not lost.")
        df = df.copy()
        df["Session"] = df["Session"].fillna(UNSPECIFIED_SESSION)

    show_preview(df)
    fingerprint = pd.util.hash_pandas_object(df, index=True).sum()
    fingerprint = f"{source}:{fingerprint}"
    if st.session_state.get("_evaluation_input_fingerprint") != fingerprint:
        st.session_state["_evaluation_input_fingerprint"] = fingerprint
        st.session_state.pop("prog_zip", None)
        st.session_state.pop("master_out", None)
    return df

def signoff_form() -> Signatory:
    """Reads Name / Position / Sign from signoff.xlsx (next to the app), or from an Excel uploaded here."""
    signoff_fingerprint = ""
    with st.expander("✍️ Sign-off (Name, Position, Sign)"):
        up = st.file_uploader("Sign-off Excel (optional: overrides signoff.xlsx for this session)",
                              type=["xlsx"], key="so_file")
        people = []
        try:
            if up is not None:
                people = read_signoffs(io.BytesIO(up.getvalue()))
                signoff_fingerprint = f"upload:{hash(up.getvalue())}"
                st.caption("Using the uploaded sign-off file.")
            elif os.path.exists(SIGNOFF_FILE):
                people = read_signoffs(SIGNOFF_FILE)
                signoff_fingerprint = f"file:{os.path.getmtime(SIGNOFF_FILE)}"
                st.caption(f"Using {SIGNOFF_FILE}.")
            else:
                st.caption(f"No {SIGNOFF_FILE} found next to the app, so reports will have blank lines to sign by hand.")
        except Exception as e:
            st.error(f"Could not read the sign-off file: {e}")

        st.download_button("📄 Download blank sign-off template", signoff_template_bytes(),
                           file_name="signoff.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    chosen = Signatory()
    if people:
        people_key = tuple((p.name, p.title) for p in people)
        if st.session_state.get("_so_people_key") != people_key:
            st.session_state.pop("so_pick", None)
            st.session_state["_so_people_key"] = people_key
        labels = {i: f"{p.name} — {p.title}".strip(" —") for i, p in enumerate(people)}
        labels[-1] = "No one (leave blank to sign by hand)"
        pick = st.selectbox("Select the signatory for generated reports", [-1] + list(range(len(people))),
                            index=1, format_func=lambda i: labels[i], key="so_pick")
        if pick >= 0:
            chosen = people[pick]
            how = "signature image" if chosen.signature else ("typed signature" if chosen.sign_text else "blank line")
            st.caption(f"Sign: {how}")
            signoff_fingerprint += f":{pick}"
        else:
            signoff_fingerprint += ":none"
    else:
        st.caption("No signatories are available. Upload a sign-off Excel file or add entries to signoff.xlsx.")
    if st.session_state.get("_signoff_fingerprint") != signoff_fingerprint:
        st.session_state["_signoff_fingerprint"] = signoff_fingerprint
        st.session_state.pop("prog_zip", None)
        st.session_state.pop("master_out", None)
    return chosen

def _count_input(label, key):
    """Headcount box that starts empty; returns None until the user fills it in."""
    val = st.number_input(label, min_value=1, value=None, step=1, key=key)
    return int(val) if val is not None else None

def single_programme_mode(df, param_cols, signatory=None):
    programmes = df["Program Name"].dropna().unique().tolist()   # order of the sheet
    selected = st.multiselect("Select Programmes:", options=programmes, default=[])
    single_input_key = repr(selected)
    if st.session_state.get("_single_selection_key") != single_input_key:
        st.session_state["_single_selection_key"] = single_input_key
        st.session_state.pop("prog_zip", None)
    if not selected:
        st.warning("Please select at least one programme.")
        return
    selected = [p for p in programmes if p in selected]   # keep sheet order, not click order

    st.markdown("### 📝 Enter Participant Headcount Setup")
    st.caption("Leave a box empty if you don't know the headcount. The report will then show "
               "'Not provided' and skip the response rate.")

    participants = {}
    totals = {}

    for prog in selected:
        prog_df = df[df["Program Name"] == prog]
        fac_rows = prog_df[prog_df["Faculty Name"].notna()]
        dept_rows = prog_df[prog_df["Faculty Name"].isna()]
        fac_counts = fac_rows.groupby(["Session", "Faculty Name"], sort=False).size()
        dept_counts = dept_rows.groupby("Session", sort=False).size()

        st.markdown(f"--- \n#### {prog}")

        # Make it obvious which rows are being treated as department sessions
        if not dept_rows.empty:
            names = [str(s) for s in dept_counts.index]
            shown = ", ".join(names[:8]) + (" …" if len(names) > 8 else "")
            if fac_rows.empty:
                st.caption(f"ℹ️ This programme has only department sessions ({len(dept_rows)} rows, no Faculty Name): {shown}.")
            else:
                st.caption(f"ℹ️ {len(dept_rows)} of {len(prog_df)} rows have no Faculty Name and are treated as "
                           f"department sessions: {shown}.")

        same_count = st.checkbox(
            f"Same total participant count for all sessions in **{prog}**",
            value=True,
            key=f"same_count_{prog}"
        )

        if same_count:
            single_val = _count_input(
                f"Participant count for all sessions & overall total in **{prog}**",
                f"single_val_{prog}")
            totals[prog] = single_val
            for (session, faculty), _ in fac_counts.items():
                participants[f"{prog} -- {faculty} -- {session}"] = single_val
            for session in dept_counts.index:
                participants[f"{prog} -- Dept -- {session}"] = single_val
        else:
            st.markdown("Enter individual class sizes per session:")
            session_vals = []
            for (session, faculty), n in fac_counts.items():
                val = _count_input(
                    f"Participants for **{session}** ({faculty}) — [{n} responses in file]",
                    f"{prog} -- {faculty} -- {session}")
                if val is not None and val < n:
                    st.warning(f"{session} ({faculty}): {n} responses but only {val} participants entered.")
                participants[f"{prog} -- {faculty} -- {session}"] = val
                session_vals.append(val)
            for session, n in dept_counts.items():
                val = _count_input(
                    f"Participants for **{session}** (Department) — [{n} responses in file]",
                    f"{prog} -- Dept -- {session}")
                if val is not None and val < n:
                    st.warning(f"{session} (Department): {n} responses but only {val} participants entered.")
                participants[f"{prog} -- Dept -- {session}"] = val
                session_vals.append(val)

            tot_val = _count_input(
                f"Total Overall Participants for **{prog}** (empty = sum of the sessions above)",
                f"total_{prog}")
            if tot_val is None and session_vals and all(v is not None for v in session_vals):
                tot_val = sum(session_vals)
            totals[prog] = tot_val

        highest = _max_session_responses(prog_df)
        if totals[prog] and highest > totals[prog]:
            st.warning(f"{prog}: the busiest session has {highest} responses but only {totals[prog]} "
                       f"participants were entered. The response rate will be shown as 100%.")

    missing = [p for p in selected if not totals.get(p)]
    if missing:
        st.info("Participants not provided for: " + ", ".join(missing) +
                ". Their overall report will show the response rate as not available.")

    invalid_headcounts = []
    for prog in selected:
        prog_df = df[df["Program Name"] == prog]
        highest = _max_session_responses(prog_df)
        total = totals.get(prog)
        if total is not None and total < highest:
            invalid_headcounts.append(f"{prog}: {highest} responses but {total} participants")
    if invalid_headcounts:
        st.error("Participant count is too low:\n\n- " + "\n- ".join(invalid_headcounts))

    st.markdown("---")
    gen_overall = st.checkbox("Generate Overall Programme Report", value=True)
    gen_individual = st.checkbox("Generate Individual Faculty Reports (ZIP)", value=True)
    combine_faculty_sessions = st.checkbox(
        "Combine each faculty's sessions into one report",
        value=True,
        disabled=not gen_individual,
        help="Each session stays in its own section. Uncheck to create a separate file for each session.",
    )
    add_sign = st.checkbox("Add sign-off block at the end (Name / Title / Sign)", value=True)
    prog_options_key = repr((
        selected, participants, totals, gen_overall, gen_individual, combine_faculty_sessions, add_sign))
    if st.session_state.get("_prog_options_key") != prog_options_key:
        st.session_state["_prog_options_key"] = prog_options_key
        st.session_state.pop("prog_zip", None)

    if st.button("🚀 Generate Programme Reports", type="primary"):
        if not (gen_overall or gen_individual):
            st.warning("Tick at least one report type.")
        elif invalid_headcounts:
            st.warning("Correct the participant counts before generating reports.")
        else:
            try:
                bar = st.progress(0.0, text="Starting…")
                data = generate_programme_zip(
                    df, selected, participants, totals, gen_overall, gen_individual, param_cols,
                    on_progress=lambda frac, name: bar.progress(frac, text=f"Finished {name}"),
                    signatory=signatory, include_signoff=add_sign,
                    combine_faculty_sessions=combine_faculty_sessions,
                )
                bar.empty()
                st.session_state["prog_zip"] = (data, len(selected))
            except Exception as exc:
                st.error(f"Report generation failed: {exc}")

    # Result survives the rerun triggered by clicking download
    if "prog_zip" in st.session_state:
        data, n = st.session_state["prog_zip"]
        st.success(f"Generated reports for {n} programme(s) successfully!")
        st.download_button("📥 Download All Reports ZIP", data, file_name="Multiple_Programmes_Reports.zip",
                           mime="application/zip")

def master_mode(df, param_cols, signatory=None):
    st.markdown("### 🏢 Master Departmental Report")
    st.markdown("A short summary for department directors and heads.")

    has_dept = _has_dept(df)
    if has_dept:
        departments = df[DEPT_COL].dropna().unique().tolist()   # order of the sheet
        sel_depts = st.multiselect("Departments:", departments, default=departments)
        missing = int(df[DEPT_COL].isna().sum())
        if missing:
            st.caption(f"{missing} rows have no Department and are left out of this report.")
        scoped = df[df[DEPT_COL].isin(sel_depts)]
    else:
        st.info(f"No '{DEPT_COL}' column found, so programmes can't be grouped by department.")
        sel_depts, scoped = [], df

    programmes = scoped["Program Name"].dropna().unique().tolist()
    selected = st.multiselect("Programmes:", programmes, default=programmes)
    selected = [p for p in programmes if p in selected]
    master_input_key = repr((sel_depts if has_dept else "all", selected))
    if st.session_state.get("_master_selection_key") != master_input_key:
        st.session_state["_master_selection_key"] = master_input_key
        st.session_state.pop("master_out", None)

    c1, c2 = st.columns(2)
    rank_n = c1.number_input("Sessions in top / bottom lists", min_value=1, max_value=10, value=3)
    max_comments = c2.number_input("Max comments per session (0 = all)", min_value=0, max_value=50, value=5)
    include_support = st.checkbox("Include support services scores (rows with no faculty name)", value=True)
    include_chart = st.checkbox("Include parameter chart", value=True)
    add_sign = st.checkbox("Add sign-off block at the end (Name / Title / Sign)", value=False)
    per_dept = (st.checkbox("One report per department (ZIP)", value=False)
                if has_dept and len(sel_depts) > 1 else False)
    master_options_key = repr((selected, sel_depts, rank_n, max_comments,
                               include_support, include_chart, add_sign, per_dept))
    if st.session_state.get("_master_options_key") != master_options_key:
        st.session_state["_master_options_key"] = master_options_key
        st.session_state.pop("master_out", None)

    if not selected:
        st.warning("Please select at least one programme.")
        return

    if st.button("🚀 Generate Master Departmental Report", type="primary"):
        sub = scoped[scoped["Program Name"].isin(selected)]
        opts = dict(rank_n=int(rank_n), max_comments=int(max_comments) or None,
                    include_support=include_support, include_chart=include_chart,
                    signatory=signatory if add_sign else None)
        try:
            with st.spinner("Compiling report..."):
                if per_dept:
                    buf = io.BytesIO()
                    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        used = set()
                        for d in sub[DEPT_COL].dropna().unique():   # sheet order
                            data = build_master_report(sub[sub[DEPT_COL] == d], param_cols,
                                                       scope_label=f"Department: {d}", **opts)
                            zf.writestr(_unique(f"{safe_name(d)}_Master_Report.docx", used), data)
                    st.session_state["master_out"] = (buf.getvalue(), "zip")
                else:
                    label = (f"Department: {', '.join(sel_depts)}"
                             if 0 < len(sel_depts) <= 3 else None)
                    st.session_state["master_out"] = (
                        build_master_report(sub, param_cols, scope_label=label, **opts), "docx")
        except Exception as exc:
            st.error(f"Master report generation failed: {exc}")

    if "master_out" in st.session_state:
        data, kind = st.session_state["master_out"]
        st.success("Report is ready.")
        if kind == "zip":
            st.download_button("📥 Download Department Reports (ZIP)", data,
                               file_name="Master_Departmental_Reports.zip", mime="application/zip")
        else:
            st.download_button(
                "📥 Download Master Departmental Report (.docx)", data,
                file_name="Master_Departmental_Evaluation_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

def main():
    st.set_page_config(page_title="SBS Evaluation Automation", layout="centered", page_icon="📊")
    st.title("📊 Strathmore Business School - Evaluation Automation")
    st.markdown("Generate Individual Faculty Reports, Programme Reports, and Periodic Departmental Reports.")

    df = ingest()
    if df is None:
        return

    param_cols = get_evaluation_params(df)
    if not param_cols:
        st.error("No evaluation parameters found in the data. Please ensure there are numeric columns for scoring.")
        return

    signatory = signoff_form()

    st.markdown("---")
    scope = st.selectbox("Select Report Generation Mode:", [
        "Single Programme Mode (Individual Faculty & Programme Reports)",
        "Master Departmental Report Mode (Cross-Programme Roll-up)",
    ])

    if "Single Programme Mode" in scope:
        single_programme_mode(df, param_cols, signatory)
    else:
        master_mode(df, param_cols, signatory)

if __name__ == "__main__":
    main()
