# Weekly Summary – Scan Files Update (Power Automate Flow)

This flow generates a weekly email summarizing scanned files for preventive and corrective maintenance work orders. It runs an Office Script on a SharePoint-hosted Excel file to extract last week’s data (Fri → Thu), formats the results into an HTML table, and sends the summary to stakeholders. If no data is found, the flow sends a short “no data” notification.

---

## 📌 Overview

- **Schedule:** Every Wednesday (UTC+07:00, Bangkok)
- **Systems:** Power Automate, SharePoint (Documents), Excel Online (Office Scripts), Outlook
- **Objectives**
  - Automate weekly summary of scanned files
  - Standardize email content with an HTML template
  - Clearly show included sheets and key fields
  - Send “no data” notification when appropriate

---

## 🧭 Flow Steps (High Level)

1) **Recurrence** → runs weekly on Wednesday  
2) **Run script (Excel Online)** → `ReadSheets_LastFriToThu` against `SCAN RECORD.xlsx`  
3) **Condition:** `rows > 0 ?`
   - **Yes (rows > 0):**  
     a. Create **HTML table** from `result.rows`  
     b. Build **IncludedSheetsString** (comma‑separated)  
     c. Build **EmailBody** (HTML template)  
     d. **Send email** (weekly summary)
   - **No (rows = 0):**  
     a. **Send email** (no‑data notification)

---

## 🧱 Flow Diagram (ASCII)
