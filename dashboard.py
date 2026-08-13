"""
Dashboard — reads submissions from the master Excel workbook (submissions.xlsx)
and shows everything: tracker, metric cards, RAG donut, Accelerate progress,
MI maturity, plus a button to download the whole master Excel.
"""

import statistics
import streamlit as st

import common
import excel_store as store

try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False


def _fmt(n):
    if n is None:
        return "—"
    a = abs(n)
    if a >= 1e9:
        return "%.2fBn" % (n / 1e9)
    if a >= 1e6:
        return "%.1fMn" % (n / 1e6)
    if a >= 1e3:
        return "%.0fk" % (n / 1e3)
    return "%.0f" % n


def _gauge(pct, title):
    if HAS_PLOTLY:
        fig = go.Figure(go.Indicator(
            mode="gauge+number", value=round(pct, 1), number={"suffix": "%"},
            title={"text": title, "font": {"size": 14}},
            gauge={"axis": {"range": [0, 100]}, "bar": {"color": "#0B0B0B"},
                   "steps": [{"range": [0, 40], "color": "#FDECEC"},
                             {"range": [40, 75], "color": "#FFF6E0"},
                             {"range": [75, 100], "color": "#E8F5E9"}],
                   "threshold": {"line": {"color": "#FFCB05", "width": 4}, "value": pct}}))
        fig.update_layout(height=240, margin=dict(l=20, r=20, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.metric(title, "%.1f%%" % pct)
        st.progress(min(1.0, pct / 100.0))


def _rag_donut(counts):
    labels = [k for k in ["Green", "Amber", "Red", "Blue", "Not set"] if counts.get(k)]
    values = [counts[k] for k in labels]
    if not values:
        st.info("No RAG data yet.")
        return
    if HAS_PLOTLY:
        fig = go.Figure(go.Pie(labels=labels, values=values, hole=0.55,
                               marker=dict(colors=[common.RAG_COLOR[k] for k in labels]), sort=False))
        fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h"))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart({"count": {k: counts[k] for k in labels}})


def _progress_row(label, pct, right):
    c1, c2, c3 = st.columns([5, 3, 2])
    c1.write(label)
    c2.progress(min(1.0, pct / 100.0))
    c3.write("**%.0f%%**  " % pct + (right or ""))


def render_dashboard(DATA):
    st.header("📊 MI & Accelerate — Executive Dashboard")

    subs = store.load_submissions()
    all_opcos = list(DATA.keys())
    submitted_opcos = sorted({p.get("opco") for p in subs})

    done, total = len(submitted_opcos), len(all_opcos)
    tc1, tc2 = st.columns([1, 3])
    tc1.metric("Submissions received", "%d / %d" % (done, total))
    missing = [o for o in all_opcos if o not in submitted_opcos]
    with tc2:
        if submitted_opcos:
            st.write("**In:** " + ", ".join(submitted_opcos))
        st.write("**Still waiting on:** " + (", ".join(missing) if missing else "— none, all in! 🎉"))

    st.download_button("⬇️ Download the master Excel (all submissions)",
                       data=store.master_bytes(), file_name="MI_Accelerate_submissions.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    use_baseline = not subs
    if use_baseline:
        st.warning("No submissions in the workbook yet — showing **baseline (deck) data** so you can "
                   "preview. As OpCos submit, the dashboard fills in automatically.")
        records = []
        for c in DATA:
            records += common.iter_baseline_records(DATA, c)
    else:
        records = common.submissions_to_records(subs)

    st.divider()
    opcos_in_data = sorted({r["opco"] for r in records})
    sel = st.multiselect("Filter OpCo(s)", opcos_in_data, default=opcos_in_data)
    records = [r for r in records if r["opco"] in sel]
    if not records:
        st.info("Nothing matches the current filter.")
        return

    accel = [r for r in records if r["kind"] == "Accelerate"]
    mi = [r for r in records if r["kind"] == "MI"]

    accel_prog = []
    for r in accel:
        pct, act, est, _ek, _ak = common.accel_progress(r["fields"])
        if pct is not None:
            accel_prog.append((r, pct, act, est))

    rag_counts = {}
    for r in records:
        if r["kind"] in ("Accelerate", "MI"):
            b = common.rag_bucket(r["fields"].get("RAG Status", ""))
            rag_counts[b] = rag_counts.get(b, 0) + 1
    avg_accel = statistics.mean([p for _, p, _, _ in accel_prog]) if accel_prog else 0.0
    rated = sum(v for k, v in rag_counts.items() if k != "Not set")
    green_pct = (rag_counts.get("Green", 0) / rated * 100.0) if rated else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("OpCos in view", len(sel))
    m2.metric("Initiatives", len(records))
    m3.metric("Avg Accelerate completion", "%.1f%%" % avg_accel)
    m4.metric("RAG green", "%.0f%%" % green_pct)

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("RAG status mix")
        _rag_donut(rag_counts)
    with c2:
        st.subheader("Overall Accelerate completion")
        _gauge(avg_accel, "Actual ÷ Estimated impact")
        st.caption("Across %d Accelerate initiative(s) with figures." % len(accel_prog))

    st.divider()
    st.subheader("🚀 Accelerate — how far each initiative is")
    if not accel_prog:
        st.info("No Accelerate initiatives with both estimated and actual impact in this selection.")
    else:
        rows = sorted(accel_prog, key=lambda x: x[1], reverse=True)
        if HAS_PLOTLY:
            labels = [("%s · %s" % (r["opco"], r["initiative"]))[:60] for r, _, _, _ in rows]
            vals = [round(p, 1) for _, p, _, _ in rows]
            colors = ["#2E7D32" if v >= 75 else "#E8A317" if v >= 40 else "#C62828" for v in vals]
            fig = go.Figure(go.Bar(x=vals, y=labels, orientation="h", marker=dict(color=colors),
                                   text=[f"{v}%" for v in vals], textposition="auto"))
            fig.update_layout(height=max(300, 26 * len(vals)), margin=dict(l=10, r=10, t=10, b=10),
                              xaxis=dict(title="% of estimated impact achieved", range=[0, 100]),
                              yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)
        by_opco = {}
        for r, p, act, est in accel_prog:
            by_opco.setdefault(r["opco"], []).append((r, p, act, est))
        for opco in sorted(by_opco):
            with st.expander("%s — %d initiative(s)" % (opco, len(by_opco[opco])),
                             expanded=(len(by_opco) <= 3)):
                for r, p, act, est in sorted(by_opco[opco], key=lambda x: x[1], reverse=True):
                    _progress_row(r["initiative"], p, "(%s / %s)" % (_fmt(act), _fmt(est)))

    st.divider()
    st.subheader("📈 MI — maturity toward target")
    mi_prog = []
    for r in mi:
        mp, cur, tgt = common.maturity_progress(r["fields"])
        if mp is not None:
            mi_prog.append((r, mp, cur, tgt))
    if not mi_prog:
        st.info("No MI initiatives with current & target maturity in this selection.")
    else:
        avg_mi = statistics.mean([p for _, p, _, _ in mi_prog])
        cA, cB = st.columns([1, 2])
        with cA:
            _gauge(avg_mi, "Avg maturity vs target")
        with cB:
            if HAS_PLOTLY:
                s = sorted(mi_prog, key=lambda x: x[1])
                labels = [("%s · %s" % (r["opco"], r["initiative"]))[:55] for r, _, _, _ in s]
                vals = [round(p, 0) for _, p, _, _ in s]
                fig = go.Figure(go.Bar(x=vals, y=labels, orientation="h", marker=dict(color="#1565C0"),
                                       text=[f"{int(v)}%" for v in vals], textposition="auto"))
                fig.update_layout(height=max(280, 24 * len(vals)), margin=dict(l=10, r=10, t=10, b=10),
                                  xaxis=dict(title="Current maturity as % of target", range=[0, 100]))
                st.plotly_chart(fig, use_container_width=True)
            else:
                for r, p, cur, tgt in sorted(mi_prog, key=lambda x: x[1], reverse=True):
                    _progress_row("%s · %s" % (r["opco"], r["initiative"]), p, "%s→%s" % (cur, tgt))

    st.caption("Data source: master Excel workbook `submissions.xlsx` (one sheet per OpCo + Overview). "
               "Accelerate % = actual ÷ estimated impact; maturity % maps Basic→Intermediate→Control→Advanced.")
