# Saving the master file to SharePoint (secure)

This makes the consolidated master workbook live in **your SharePoint**, written by the app
with a service credential. The OpCo submitters use the form but **cannot open, download or
edit the master file** — only Group EBU can.

## What you get
- **Master workbook** `submissions.xlsx` stored in SharePoint (you choose the folder).
- Every entry is **date & time-stamped (UTC)** in three columns:
  `Submitted (UTC)`, `Submitted date`, `Submitted time`.
- A hidden **`_audit`** sheet logs *every* submit event (created / updated) with timestamp,
  OpCo, person and email — a full history even when an OpCo re-submits.
- The **Dashboard** (and the master download) is protected by an **admin passcode**.
- Submitters only ever get a copy of **their own** submission.

---

## Step 1 — Register an Azure AD app (one-time; needs MTN IT)
1. Azure Portal → **App registrations** → **New registration** → name it e.g. *MI-Accelerate-App*.
2. After it's created, note the **Application (client) ID** and **Directory (tenant) ID**.
3. **Certificates & secrets** → **New client secret** → copy the **Value** (not the ID).
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
   → add **Sites.ReadWrite.All** → then **Grant admin consent** (IT does this).

> Application permission + admin consent is what lets the app write to SharePoint without any
> user signing in — and is exactly why submitters can't reach the file.

## Step 2 — Pick where the file lives
Decide the SharePoint site and path, e.g.
`https://mtncloud.sharepoint.com/sites/EBU-Strategy` → folder `MI and Accelerate`.
- `site_hostname` = `mtncloud.sharepoint.com`
- `site_path`     = `/sites/EBU-Strategy`
- `file_path`     = `MI and Accelerate/submissions.xlsx`  (the app creates it on first submit)

## Step 3 — Add the secrets
Streamlit Cloud → your app → **Settings → Secrets**, paste (see `secrets.toml.example`):

```toml
admin_pin = "choose-a-strong-passcode"

[sharepoint]
tenant_id     = "...."
client_id     = "...."
client_secret = "...."
site_hostname = "mtncloud.sharepoint.com"
site_path     = "/sites/EBU-Strategy"
file_path     = "MI and Accelerate/submissions.xlsx"
```

Save. Done — the sidebar will now read *“saved securely to Group EBU (SharePoint · …)”*.

---

## How access is locked down
| Who | Can submit their OpCo | Can see Dashboard | Can open master file |
|-----|:---:|:---:|:---:|
| OpCo submitters (the 20) | ✅ | ❌ (needs passcode) | ❌ |
| You / Group EBU (passcode) | ✅ | ✅ | ✅ (in SharePoint, or via Dashboard) |

- Submitters never receive the file or the credentials — the app writes to SharePoint on the
  server side using the app registration.
- Keep the SharePoint **folder permissions** restricted to Group EBU as a second layer.

## If IT can't approve the app registration
Tell me — I'll swap `excel_store.py` for a **Power Automate** version that appends each entry
to the SharePoint Excel using *your own* OneDrive/SharePoint connection (no IT approval, no
credentials in the app). Same outcome: master lives in SharePoint, submitters can't access it.

## No secrets set?
The app falls back to a **local `submissions.xlsx`** next to the code — perfect for testing on
your PC. Timestamps and the audit log work the same way.
