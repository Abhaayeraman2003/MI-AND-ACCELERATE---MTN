# DGW MI & Accelerate — Excel edition (all 14 OpCos)

Everything is stored in ONE master Excel workbook: **submissions.xlsx**.

## Run it
Windows: double-click `run_windows.bat`  ·  Mac/Linux: `./run_mac_linux.sh`
…or:  `pip install -r requirements.txt`  then  `streamlit run streamlit_app.py`

## How it works
- **📝 Submit** → the OpCo's answers are written into `submissions.xlsx`
  (sheet "Submissions" = clean filterable table; one row per initiative with OpCo, RAG,
  Accelerate %, Actual, Estimated, Maturity, Comment). Re-submitting the same OpCo + month
  overwrites its rows (no duplicates). Each person also gets an Excel copy to download.
- **📊 Dashboard** → reads `submissions.xlsx`, shows the tracker + charts + RAG mix +
  Accelerate progress, with a "Download master Excel" button.

## Files (push ALL of these to GitHub)
streamlit_app.py · dashboard.py · excel_store.py · common.py · excel_builder.py ·
data.json (all 14 OpCos) · requirements.txt · .streamlit/config.toml

⚠️ Do NOT keep an old `db.py` in the repo — this version uses Excel, not a database.

## ⚠️ Streamlit Cloud note
Streamlit Cloud storage is temporary — `submissions.xlsx` is wiped on restart and viewers
may hit different servers. Great for a **demo**; for real 20-person collection run it
**locally** or host the master workbook on **OneDrive/SharePoint** (only excel_store.py changes).
