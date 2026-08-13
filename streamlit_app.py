"""
DGW MI & Accelerate — Excel-backed edition.

  📝 Submit update — OpCos edit what changed; on submit the row(s) are written into ONE
     master Excel workbook (submissions.xlsx). An individual Excel copy is also offered.
  📊 Dashboard     — reads the master workbook and shows everything, and lets you download
     the whole master Excel.

Everything lives in submissions.xlsx. Only excel_store.py touches the file, so moving the
master workbook to OneDrive/SharePoint later means changing just that one file.
"""

import os
import json
import datetime
import streamlit as st

import common
import excel_store as store
from common import (
    MONTHS, KIND_LABEL, KIND_ORDER, READONLY_COLS, RAG_OPTS, MATURITY,
    ascii_slug, is_section, field_type, rag_normalise, row_info, col_label,
)
from excel_builder import build_workbook
from dashboard import render_dashboard

st.set_page_config(page_title="MTN EBU · MI & Accelerate", page_icon="🟡", layout="wide")


@st.cache_data
def load_data():
    with open(os.path.join(os.path.dirname(__file__), "data.json"), "r", encoding="utf-8") as fh:
        return json.load(fh)


DATA = load_data()

st.markdown(
    """
    <div style="background:#0B0B0B;padding:14px 20px;border-radius:10px;display:flex;align-items:center;gap:14px;margin-bottom:8px">
      <div style="width:36px;height:36px;border-radius:50%;background:#FFCB05;display:flex;align-items:center;justify-content:center;font-weight:800;color:#000">MTN</div>
      <div>
        <div style="color:#fff;font-size:17px;font-weight:600">Accelerate &amp; Maturity Index — Monthly OpCo Update</div>
        <div style="color:#B9B9B9;font-size:12px">MTN Business Africa · Group EBU · Strategy &amp; Transformation</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

view = st.sidebar.radio("View", ["📝 Submit update", "📊 Dashboard"])
st.sidebar.divider()

if view == "📊 Dashboard":
    render_dashboard(DATA)
    st.stop()


def build_payload(country, reporting_month_name, year, name, email, answers):
    out = {"opco": country, "reportingMonth": "%s %s" % (reporting_month_name, year),
           "submittedBy": name, "email": email,
           "submittedAt": datetime.datetime.utcnow().isoformat() + "Z", "sections": []}
    blocks = sorted(DATA[country], key=lambda b: (KIND_ORDER.index(b["kind"]), b["slide"]))
    for b in blocks:
        kind = b["kind"]
        ro = READONLY_COLS.get(kind, 2)
        items = []
        group = ""
        for ri, row in enumerate(b["rows"]):
            if is_section(row):
                group = next((c for c in row if c and c.strip()), "")
                continue
            name_i, group, start = row_info(b, row, group)
            if not name_i:
                continue
            rec = {"section": "" if group == name_i else group, "initiative": name_i, "fields": {}, "changed": False}
            for i in range(start, ro):
                if i < len(row) and row[i]:
                    rec["fields"][b["headers"][i] if i < len(b["headers"]) else ("col%d" % i)] = row[i]
            for i in range(ro, len(b["headers"])):
                h_raw = b["headers"][i]
                if not h_raw or not h_raw.strip():
                    continue
                h = col_label(b, i, reporting_month_name)
                key = "|".join([kind, str(b["slide"]), str(ri), str(i)])
                prev = ((row[i] if i < len(row) else "") or "").strip()
                if field_type(h_raw) == "rag":
                    prev = rag_normalise(prev)
                val = answers.get(key, prev)
                rec["fields"][h] = val
                if val != prev:
                    rec["changed"] = True
            ckey = "|".join([kind, str(b["slide"]), str(ri), "comment"])
            c = answers.get(ckey, "")
            if c and c.strip():
                rec["fields"]["Comment / risk"] = c.strip()
                rec["changed"] = True
            items.append(rec)
        out["sections"].append({"type": KIND_LABEL[kind], "sourceSlide": b["slide"], "items": items})
    out["itemsUpdated"] = sum(len([i for i in s["items"] if i["changed"]]) for s in out["sections"])
    return out


with st.sidebar:
    st.subheader("Your details")
    country = st.selectbox("Select your OpCo", ["— Choose your country —"] + sorted(DATA.keys()))
    today = datetime.date.today()
    rmonth = st.selectbox("Reporting month", MONTHS, index=today.month - 1)
    ryear = st.selectbox("Year", [str(today.year - 1), str(today.year), str(today.year + 1)], index=1)
    uname = st.text_input("Your name", placeholder="Full name")
    uemail = st.text_input("Your email", placeholder="name.surname@mtn.com")
    st.caption("Saved to master Excel: `submissions.xlsx` ✅")

if country not in DATA:
    st.info("👈 Choose your OpCo in the sidebar to load your initiatives, or switch to the **Dashboard** view.")
    st.stop()

valid = bool(uname.strip()) and bool(uemail.strip()) and "@" in uemail

blocks = sorted(DATA[country], key=lambda b: (KIND_ORDER.index(b["kind"]), b["slide"]))
counts = {}
for b in blocks:
    n = len([r for r in b["rows"] if not is_section(r)])
    counts[b["kind"]] = counts.get(b["kind"], 0) + n
st.success("Loaded for **%s** — " % country +
           " · ".join("%d %s" % (counts[k], KIND_LABEL[k].lower()) for k in KIND_ORDER if counts.get(k)))
if not valid:
    st.warning("Enter your name and a valid email in the sidebar to enable submission.")

answers = {}
present_kinds = [k for k in KIND_ORDER if any(b["kind"] == k for b in blocks)]
tabs = st.tabs([KIND_LABEL[k] for k in present_kinds])
for tab, kind in zip(tabs, present_kinds):
    with tab:
        st.caption({
            "MI": "Maturity Index initiatives. Update activity, due dates, maturity and RAG.",
            "Accelerate": "Accelerate initiatives. Update progress, estimated impact and YTD performance.",
            "Priorities": "Monthly KPI values — enter this month's actuals.",
            "MI Tracker": "Actions taken this month against each initiative, plus RAG and timeline.",
        }.get(kind, ""))
        ro = READONLY_COLS.get(kind, 2)
        for b in [b for b in blocks if b["kind"] == kind]:
            group = ""
            for ri, row in enumerate(b["rows"]):
                if is_section(row):
                    group = next((c for c in row if c and c.strip()), "")
                    st.markdown("###### 🟨 %s" % group)
                    continue
                name_i, group, start = row_info(b, row, group)
                if not name_i:
                    continue
                with st.expander(name_i, expanded=False):
                    if group and group != name_i:
                        st.markdown("_%s_" % group)
                    ctx = []
                    for i in range(start, ro):
                        if i < len(row) and row[i] and row[i].strip():
                            ctx.append("**%s:** %s" % (b["headers"][i] if i < len(b["headers"]) else "", row[i]))
                    if ctx:
                        st.caption("  \n".join(ctx))
                    cols = st.columns(2)
                    ci = 0
                    for i in range(ro, len(b["headers"])):
                        h_raw = b["headers"][i]
                        if not h_raw or not h_raw.strip():
                            continue
                        h = col_label(b, i, rmonth)
                        key = "|".join([kind, str(b["slide"]), str(ri), str(i)])
                        prev = ((row[i] if i < len(row) else "") or "").strip()
                        ft = field_type(h_raw)
                        if ft == "rag":
                            pv = rag_normalise(prev)
                            with cols[ci % 2]:
                                answers[key] = st.selectbox(h, RAG_OPTS, index=RAG_OPTS.index(pv) if pv in RAG_OPTS else 0, key=key)
                            ci += 1
                        elif ft == "maturity":
                            opts = [""] + MATURITY
                            with cols[ci % 2]:
                                answers[key] = st.selectbox(h, opts, index=opts.index(prev) if prev in opts else 0, key=key)
                            ci += 1
                        elif ft == "short":
                            with cols[ci % 2]:
                                answers[key] = st.text_input(h, value=prev, key=key)
                            ci += 1
                        else:
                            answers[key] = st.text_area(h, value=prev, key=key, height=80)
                    ckey = "|".join([kind, str(b["slide"]), str(ri), "comment"])
                    answers[ckey] = st.text_area("Comment / risk (optional)", key=ckey,
                                                 placeholder="Blockers, support needed, scope changes…", height=68)

st.divider()
if "submitted" not in st.session_state:
    st.session_state.submitted = None

if st.button("Submit to Group EBU ✓", type="primary", disabled=not valid):
    payload = build_payload(country, rmonth, ryear, uname.strip(), uemail.strip(), answers)
    try:
        n = store.save_submission(payload, ryear, rmonth)
        xlsx = build_workbook(payload)
        st.session_state.submitted = {
            "opco": country, "period": "%s %s" % (rmonth, ryear), "count": n,
            "updated": payload["itemsUpdated"], "xlsx": xlsx,
            "fname": "MI_Accelerate_%s_%s%02d.xlsx" % (ascii_slug(country), ryear, MONTHS.index(rmonth) + 1),
        }
    except Exception as exc:
        st.error("Could not save to the Excel workbook: %s" % exc)
        st.session_state.submitted = None

if st.session_state.submitted:
    s = st.session_state.submitted
    st.success("**Saved to the master Excel.** %s — %s · %d initiative(s) updated. "
               "The workbook now holds %d OpCo submission(s). It appears on the Dashboard."
               % (s["opco"], s["period"], s["updated"], s["count"]))
    c1, c2 = st.columns(2)
    c1.download_button("⬇️ Download your OpCo Excel", data=s["xlsx"], file_name=s["fname"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("⬇️ Download master Excel (all OpCos)", data=store.master_bytes(),
                       file_name="MI_Accelerate_submissions.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
