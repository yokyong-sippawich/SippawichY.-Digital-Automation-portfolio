# Weekly Summary – Scan Files Update (Power Automate Flow)

This flow generates a weekly email summarizing scanned files for preventive and corrective maintenance work orders. It runs an Office Script on a SharePoint-hosted Excel file to extract last week’s data (Fri → Thu), formats the results into an HTML table, composes a branded email body, and sends the summary to stakeholders.

---

## 📌 Overview

- **Schedule:** Weekly on **Wednesday** (UTC+07:00 Bangkok)
- **Systems used:** Power Automate, SharePoint (Documents), Excel Online (Office Scripts), Outlook
- **Purpose:**
  - Automate weekly summary of scanned files
  - Standardize email content with HTML template
  - Clearly show included sheets and key fields

---

## 🧩 Flowchart (Mermaid)

```mermaid
flowchart TB
  %% --------------- LANES ---------------
  subgraph PA[Power Automate]
    direction TB
    A["⏰ Recurrence<br/>Weekly (Wed)"]
    G["🟩 Condition<br/>rows > 0 ?"]
  end

  subgraph XL[Excel Online]
    direction TB
    B["🧮 Run script<br/>ReadSheets_LastFriToThu"]
  end

  subgraph TRUE[If True]
    direction TB
    C["📧 Send email (V2)<br/>No data notification"]
  end

  subgraph FALSE[If False]
    direction TB
    D["🧱 Create HTML table<br/>from result.rows"]
    E["🔤 IncludedSheetsString<br/>join(coalesce(...), ', ')"]
    F["📝 EmailBody<br/>concat(html...)"]
    H["📧 Send email (V2)<br/>Weekly summary"]
  end

  %% --------------- FLOW ---------------
  A --> B --> G
  G -->|Yes (rows > 0)| D --> E --> F --> H
  G -->|No (rows = 0)| C

