"""
Excel storage for the DGW MI & Accelerate app.

Everything is kept in ONE master workbook: submissions.xlsx
  * Sheet "Submissions" — a clean, human-friendly flat table (one row per initiative)
  * Sheet "_raw"        — hidden bookkeeping: the full JSON of each submission

Re-submitting the same OpCo + Year + Month overwrites that OpCo's rows (no duplicates).
"""

import os
import json
from openpyxl import Workbook, load_workbook

import common

MASTER_PATH = os.environ.get("MASTER_XLSX",
                             os.path.join(os.path.dirname(__file__), "submissions.xlsx"))

FLAT_HEADERS = ["OpCo", "Reporting month", "Submitted by", "Email", "Type", "Section",
                "Initiative", "RAG", "Accelerate %", "Actual", "Estimated",
                "Maturity %", "Current", "Target", "Comment / risk", "Submitted at"]


def _ensure_workbook():
    if os.path.exists(MASTER_PATH):
        return load_workbook(MASTER_PATH)
    wb = Workbook()
    ws = wb.active
    ws.title = "Submissions"
    ws.append(FLAT_HEADERS)
    raw = wb.create_sheet("_raw")
    raw.append(["opco", "year", "month", "payload_json"])
    raw.sheet_state = "hidden"
    wb.save(MASTER_PATH)
    return wb


def _read_raw(wb):
    if "_raw" not in wb.sheetnames:
        return []
    out = []
    ws = wb["_raw"]
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        opco, year, month, pj = (list(r) + [None] * 4)[:4]
        try:
            payload = json.loads(pj) if pj else {}
        except Exception:
            payload = {}
        out.append((opco, str(year), str(month), payload))
    return out


def _style_header(ws):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="0B0B0B")
        cell.fill = PatternFill("solid", fgColor="FFCB05")
    widths = [16, 16, 18, 24, 20, 20, 34, 16, 12, 14, 14, 12, 14, 14, 40, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def _rebuild_flat(wb, raws):
    from openpyxl.styles import PatternFill
    if "Submissions" in wb.sheetnames:
        del wb["Submissions"]
    ws = wb.create_sheet("Submissions", 0)
    ws.append(FLAT_HEADERS)
    RAGFILL = {"Green": "E8F5E9", "Amber": "FFF6E0", "Red": "FDECEC", "Blue": "EAF2FC"}
    for opco, year, month, payload in raws:
        for rec in common.submissions_to_records([payload]):
            f = rec["fields"]
            ap = act = est = mp = cur = tgt = ""
            if rec["kind"] == "Accelerate":
                p, a, e, _ek, _ak = common.accel_progress(f)
                ap = round(p, 1) if p is not None else ""
                act = a if a is not None else ""
                est = e if e is not None else ""
            if rec["kind"] == "MI":
                p, cu, tg = common.maturity_progress(f)
                mp = round(p, 0) if p is not None else ""
                cur, tgt = cu or "", tg or ""
            rag = common.rag_bucket(f.get("RAG Status", ""))
            ws.append([rec["opco"], payload.get("reportingMonth", ""), payload.get("submittedBy", ""),
                       payload.get("email", ""), rec["type"], rec.get("section", ""), rec["initiative"],
                       rag, ap, act, est, mp, cur, tgt, f.get("Comment / risk", ""),
                       payload.get("submittedAt", "")])
            if rag in RAGFILL:
                ws.cell(row=ws.max_row, column=8).fill = PatternFill("solid", fgColor=RAGFILL[rag])
    _style_header(ws)


def save_submission(payload, year, month):
    """Insert/replace this OpCo's submission for the period. Returns OpCo count in master."""
    wb = _ensure_workbook()
    raws = _read_raw(wb)
    opco = payload.get("opco")
    # upsert: drop any existing row for this OpCo+period, then add the new one
    raws = [(o, y, m, p) for (o, y, m, p) in raws
            if not (o == opco and str(y) == str(year) and str(m) == str(month))]
    raws.append((opco, str(year), str(month), payload))

    # rewrite the hidden _raw sheet
    if "_raw" in wb.sheetnames:
        del wb["_raw"]
    raw = wb.create_sheet("_raw")
    raw.append(["opco", "year", "month", "payload_json"])
    for o, y, m, p in raws:
        raw.append([o, y, m, json.dumps(p, ensure_ascii=False)])
    raw.sheet_state = "hidden"

    # rebuild the human-friendly flat sheet and SAVE to disk
    _rebuild_flat(wb, raws)
    wb.save(MASTER_PATH)
    return len(raws)


def load_submissions():
    if not os.path.exists(MASTER_PATH):
        return []
    wb = load_workbook(MASTER_PATH, data_only=True)
    return [p for (_o, _y, _m, p) in _read_raw(wb) if p]


def master_bytes():
    _ensure_workbook()
    with open(MASTER_PATH, "rb") as fh:
        return fh.read()


def list_meta():
    wb = _ensure_workbook()
    out = []
    for opco, year, month, p in _read_raw(wb):
        out.append({"opco": opco, "year": year, "month": month,
                    "submitted_by": p.get("submittedBy", ""),
                    "items_updated": p.get("itemsUpdated", ""),
                    "submitted_at": p.get("submittedAt", "")})
    return out
