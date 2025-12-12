
# Flow Execution History (XrmToolBox Plugin)

> View and analyze **Power Automate** (cloud) **Flow Execution History stored in Dataverse**.  
> Filter by **solution**, **flow**, or **date range only**; open the **run in the portal**, **copy run URL**, and **export to Excel/CSV**—all from XrmToolBox.

!Flow Execution History - screenshot

---

## ✨ Features

- **Dataverse‑only** querying of run history (`flowrun` elastic table)
- **Two ways to query**
  - **Flow-specific:** select a *Solution* → *Flow*, or type a **Flow Name**
  - **Date-only:** no selection, list **all runs** in the date range (optionally filter by selected solution)
- **Columns**: Solution, Flow, Status, Start/End, Duration (ms), Error, Run Id, Run URL
- **Deep links**:
  - **Run URL** link column → opens the run in Power Automate portal
  - **Copy Run URL** button + right‑click context menu
- **Export**: Excel (**.xlsx**) via ClosedXML, and optional **.csv**
- **Prominent status bar** with icons, color coding, and subtle flash to show progress/errors
- **Environment Id** textbox with persistence and *best‑effort* auto populate for Default environments

> **Note:** Only **solution‑aware** cloud flows have run history in Dataverse. Personal “My flows” are not included by design. Dataverse run history is typically retained for ~**28 days** depending on your admin settings.

---

## 🧩 Prerequisites

- **XrmToolBox** (latest)
- Permissions to read:
  - `flowrun` (Flow Run) elastic table (read)
  - `workflow` (flows metadata) (read)
  - `solution` / `solutioncomponent` (read)
- For Export to Excel:
  - The plugin ships with **ClosedXML** & **OpenXML** dependencies via NuGet; Excel is **not required** on the machine.

---

## 🚀 Installation (from Tool Library)

1. Open **XrmToolBox** → **Tool Library**.
2. Search for **Flow Execution History**.
3. Click **Install**.
4. Launch the tool from your **Plugins** list.

---

## 🔧 Usage

1. **Connect** XrmToolBox to your Dataverse environment.
2. (Optional) **Environment Id**:
   - Paste the **Environment Id** (e.g., `Default-<orgGuid>` or an environment GUID).  
   - You can copy it from a **Power Automate run URL** or **Admin Center**.  
   - The tool persists the value in its **Settings**. For **Default** environments, it can auto‑suggest `Default-<OrganizationId>`.
3. Choose how to query:
   - **Flow‑specific**:
     - Pick **Solution** (default shows `---Select solution---`) → pick **Flow**, or type a **Flow Name** if you don’t want to use the dropdown.
   - **Date‑only**:
     - Leave **Flow** unselected and **Flow Name** empty.
     - Set **Start / End** and click **Load Runs**.
     - If a **Solution** is selected, results will be filtered to flows in that solution.
4. **Open** a run from the **Run URL** link or use **Copy Run URL**.
5. **Export**:
   - Click **Export to Excel** to generate `.xlsx` (auto headers, filters, and table formatting).
   - (Optional) Use **Export to CSV** if enabled.

---

## 🧠 Tips

- The **status bar** shows informative icons & colors:
  - ℹ Info, ✅ Success, ⚠ Warning, ❌ Error
- If Run URL cells are empty, ensure **Environment Id** is populated.
- Use **date‑only** mode to quickly scan recent failures across many flows.

---

## ⚠ Limitations

- **Dataverse‑only**: The tool reads the **`flowrun`** elastic table; **personal (non‑solution) flows** do **not** appear.
- **Retention**: Run history is typically retained for ~**28 days** (configurable by admins). Older runs won’t be returned.
- **Inputs/Outputs**: Detailed action I/O is best viewed in the **Power Automate portal run page** (open via Run URL). This tool does not pull preview/undocumented APIs for I/O to remain stable and supportable.
