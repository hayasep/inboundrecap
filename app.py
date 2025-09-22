import os
import io
import json
import math
import smtplib
from email.message import EmailMessage
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, send_file, abort, jsonify
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from dotenv import load_dotenv

# Load .env that sits next to this file (reliable regardless of CWD)
load_dotenv(dotenv_path=Path(__file__).with_name(".env"), override=True)

# ---------------- Config ----------------
TEMPLATE_PATH = os.environ.get("EXCEL_TEMPLATE_PATH", "backstock report.xlsx")
TEMPLATE_PATH = str(Path(TEMPLATE_PATH).resolve())
SHEET_NAME = os.environ.get("SHEET_NAME", "")

SMTP_SERVER   = os.environ.get("SMTP_SERVER", "")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USERNAME = os.environ.get("SMTP_USERNAME", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
SENDER_EMAIL  = os.environ.get("SENDER_EMAIL", "")
DEFAULT_RECIPIENT = os.environ.get("DEFAULT_RECIPIENT", "")

# Optional behaviors
ALLOW_BLANK_OVERWRITE = os.environ.get("ALLOW_BLANK_OVERWRITE", "0").lower() in ("1", "true", "yes")
UNIFORM_ROW_HEIGHT_PX_ENV = os.environ.get("UNIFORM_ROW_HEIGHT_PX")  # baseline height (pixels) for all rows
DEFAULT_COL_WIDTH_PX_ENV  = os.environ.get("DEFAULT_COL_WIDTH_PX")   # optional default col width (px)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "change-me")


# ---------------- Helpers ----------------
def merged_anchor(ws, r, c):
    """Return (row,col) of the writable anchor for (r,c) if it's inside a merged range."""
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
            return mr.min_row, mr.min_col
    return r, c


def _pt_to_px(pt: float) -> int:
    return int(round(float(pt) * (96.0 / 72.0)))


def _px_to_pt(px: float) -> float:
    return float(px) * 72.0 / 96.0


def wb_to_luckysheet_json(xlsx_path, sheet_name=""):
    """
    Convert an Excel sheet into Luckysheet JSON:
      - values (.v)
      - merged ranges
      - explicit column widths
      - UNIFORM default row height for the web grid
    """
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    # ---- Decide uniform row height (px) for WEB ----
    if UNIFORM_ROW_HEIGHT_PX_ENV:
        uniform_row_px = int(UNIFORM_ROW_HEIGHT_PX_ENV)
    else:
        first_rd = ws.row_dimensions.get(1)
        if first_rd and first_rd.height:
            uniform_row_px = _pt_to_px(first_rd.height)
        else:
            drh_pt = getattr(ws.sheet_format, "defaultRowHeight", None)
            uniform_row_px = _pt_to_px(drh_pt) if drh_pt else 20

    # ---- Default column width (optional) ----
    default_col_px = None
    dcw_chars = getattr(ws.sheet_format, "defaultColWidth", None)
    if dcw_chars:
        default_col_px = int(round(float(dcw_chars) * 7))
    if DEFAULT_COL_WIDTH_PX_ENV:
        default_col_px = int(DEFAULT_COL_WIDTH_PX_ENV)

    # ---- Data (2D array of cell objects or None) ----
    data = []
    for r in range(1, max_row + 1):
        row_arr = []
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            row_arr.append({"v": v} if v not in (None, "") else None)
        data.append(row_arr)

    # ---- Merges ----
    merges = {}
    for mr in ws.merged_cells.ranges:
        r0 = mr.min_row - 1
        c0 = mr.min_col - 1
        rs = mr.max_row - mr.min_row + 1
        cs = mr.max_col - mr.min_col + 1
        merges[f"{r0}_{c0}"] = {"r": r0, "c": c0, "rs": rs, "cs": cs}

    # ---- Column widths (explicit only) ----
    columnlen = {}
    for c in range(1, max_col + 1):
        letter = get_column_letter(c)
        cd = ws.column_dimensions.get(letter)
        if cd and cd.width:
            columnlen[c - 1] = int(float(cd.width) * 7)  # chars → px heuristic

    # Row heights: keep empty — we use defaultRowHeight on the web
    rowlen = {}

    sheet = {
        "name": ws.title,
        "data": data,
        "config": {
            "merge": merges,
            "columnlen": columnlen,
            "rowlen": rowlen,
        }
    }
    return {
        "info": {"name": ws.title},
        "sheets": [sheet],
        "defaults": {
            "rowHeightPx": uniform_row_px,
            "colWidthPx": default_col_px,
        }
    }


def _uniform_row_height_pt(ws):
    """Return the baseline uniform row height in POINTS for saving to Excel."""
    px = os.environ.get("UNIFORM_ROW_HEIGHT_PX")
    if px:
        return _px_to_pt(float(px))
    rd1 = ws.row_dimensions.get(1)
    if rd1 and rd1.height:
        return float(rd1.height)
    drh = getattr(ws.sheet_format, "defaultRowHeight", None)
    if drh:
        return float(drh)
    return 15.0  # Excel-ish default


def _col_char_width(ws, col_index: int) -> float:
    """Return column width in 'characters' (Excel units), with sensible defaults."""
    letter = get_column_letter(col_index)
    cd = ws.column_dimensions.get(letter)
    if cd and cd.width:
        return float(cd.width)
    # Excel default width ≈ 8.43 chars
    return float(getattr(ws.sheet_format, "defaultColWidth", 8.43))


def _merged_span_for_cell(ws, r: int, c: int) -> int:
    """How many columns this cell spans (1 if not a merged anchor)."""
    for mr in ws.merged_cells.ranges:
        if mr.min_row == r and mr.min_col == c:
            return mr.max_col - mr.min_col + 1
    return 1


def _estimate_needed_lines(text: str, chars_per_line: int) -> int:
    if chars_per_line <= 0:
        chars_per_line = 1
    lines = 0
    parts = str(text).splitlines() or [""]
    for part in parts:
        length = len(part)
        lines += max(1, math.ceil(length / chars_per_line))
    return max(1, lines)


def apply_cells_and_export(cells):
    """Apply edited cells to a copy of the template and return (BytesIO, filename)."""
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found at {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET_NAME] if (SHEET_NAME and SHEET_NAME in wb.sheetnames) else wb.active

    # 1) Write values (only from merged anchors; skip blank wipes)
    for cell in cells:
        r = int(cell.get("r", 0) or 0)
        c = int(cell.get("c", 0) or 0)
        v = cell.get("v", "")
        if r <= 0 or c <= 0:
            continue

        tr, tc = merged_anchor(ws, r, c)
        if (r, c) != (tr, tc):
            continue

        val = v
        if isinstance(v, str):
            s = v.strip()
            if s == "":
                if not ALLOW_BLANK_OVERWRITE:
                    continue
                val = ""
            else:
                try:
                    if s.isdigit():
                        val = int(s)
                    else:
                        val = float(s)
                except Exception:
                    val = v

        try:
            ws.cell(row=tr, column=tc, value=val)
        except Exception as e:
            print(f"Write failed r{r} c{c} -> r{tr} c{tc}: {e}")

    # 2) Autosize rows for content: wrap text + increase row height as needed
    base_pt = _uniform_row_height_pt(ws)

    for r in range(1, ws.max_row + 1):
        max_lines = 1
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val is None or str(val) == "":
                continue

            # effective width (in Excel character units), accounting for merged span
            span = _merged_span_for_cell(ws, r, c)
            chars_capacity = 0.0
            for cc in range(c, c + span):
                chars_capacity += _col_char_width(ws, cc)
            chars_capacity = max(1, int(round(chars_capacity)))

            # estimate lines for this cell
            needed = _estimate_needed_lines(str(val), chars_capacity)
            if needed > 1:
                cell.alignment = Alignment(wrap_text=True)
            max_lines = max(max_lines, needed)

        ws.row_dimensions[r].height = max(base_pt, base_pt * max_lines)

    # 3) Also bump the sheet default (helps Excel render new rows similarly)
    ws.sheet_format.defaultRowHeight = base_pt

    # 4) Save to memory
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_filename = f"backstock-report-filled-{ts}.xlsx"
    return out, out_filename


def send_via_smtp(file_bytes, filename, recipient, subject, body):
    """Send the attachment via SMTP; raise on error with helpful message."""
    if not recipient:
        raise RuntimeError("No recipient provided.")
    missing = [k for k, v in {
        "SENDER_EMAIL": SENDER_EMAIL,
        "SMTP_SERVER": SMTP_SERVER,
        "SMTP_USERNAME": SMTP_USERNAME,
        "SMTP_PASSWORD": SMTP_PASSWORD,
    }.items() if not v]
    if missing:
        raise RuntimeError(f"SMTP config missing: {', '.join(missing)}")

    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient
    msg["Subject"] = subject or "Backstock Report"
    msg.set_content(body or "Please see the attached report.")
    msg.add_attachment(file_bytes,
                       maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       filename=filename)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.send_message(msg)


# ---------------- Routes ----------------
@app.route("/", methods=["GET"])
def home():
    return render_template("index.html", default_recipient=DEFAULT_RECIPIENT)


@app.route("/json", methods=["GET"])
def json_sheet():
    if not os.path.exists(TEMPLATE_PATH):
        return ("Template not found", 404)
    try:
        export_json = wb_to_luckysheet_json(TEMPLATE_PATH, SHEET_NAME)
        return app.response_class(
            response=json.dumps(export_json, default=str),
            status=200,
            mimetype="application/json",
            headers={
                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
                "Pragma": "no-cache",
                "Expires": "0",
            }
        )
    except Exception as e:
        return (f"Failed to convert workbook: {e}", 500)


@app.route("/download", methods=["POST"])
def download():
    payload = request.get_json(silent=True) or {}
    cells = payload.get("cells", [])
    try:
        out, out_filename = apply_cells_and_export(cells)
        return send_file(
            out, as_attachment=True,
            download_name=out_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return (f"Download failed: {e}", 500)


@app.route("/email", methods=["POST"])
def email():
    payload = request.get_json(silent=True) or {}
    cells = payload.get("cells", [])
    recipient = (payload.get("recipient") or "").strip()
    subject = payload.get("subject") or "Backstock Report"
    body = payload.get("body") or "Please see the attached report."
    try:
        out, out_filename = apply_cells_and_export(cells)
        send_via_smtp(out.getvalue(), out_filename, recipient, subject, body)
        return jsonify({"ok": True, "message": f"Email sent to {recipient}."})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


# Optional: quick env check
@app.route("/envcheck")
def envcheck():
    keys = ["SENDER_EMAIL", "SMTP_SERVER", "SMTP_PORT", "SMTP_USERNAME", "SMTP_PASSWORD"]
    return {k: bool(os.environ.get(k)) for k in keys}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=True)
