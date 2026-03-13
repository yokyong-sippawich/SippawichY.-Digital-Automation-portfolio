# Weekly Summary – Scan Files Update (Power Automate Flow)

This Power Automate flow generates a weekly email summarizing scanned files for preventive and corrective maintenance work orders. It runs an Office Script on a SharePoint-hosted Excel file to extract last week’s data (Fri → Thu), formats the results into an HTML table, composes a branded email body, and sends the summary to stakeholders.

---

## 📌 Overview

- Schedule: Every Wednesday (UTC+07:00, Bangkok)
- Systems used: Power Automate, SharePoint (Documents), Excel Online (Office Scripts), Outlook
- Purpose:
  - Automate weekly summary of scanned files
  - Standardize email content with an HTML template
  - Clearly show included sheets and key fields

---

## 🧩 Flowchart (Mermaid, **Styled version**)

```mermaid
flowchart TB
  %% -------- Lanes --------
  subgraph PA[Power Automate]
    direction TB
    A["Recurrence (Weekly, Wed, UTC+07)"]
    C{"rows > 0 ?"}
  end

  subgraph XL[Excel Online]
    direction TB
    B["Run script: ReadSheets_LastFriToThu"]
  end

  subgraph FX[Formatting]
    direction TB
    D["Create HTML table"]
    E["Compose IncludedSheetsString"]
    F["Compose EmailBody (HTML)"]
  end

  subgraph OL[Outlook]
    direction TB
    G["Send summary email"]
    H["Send no-data email"]
  end

  %% -------- Flow --------
  A --> B --> C
  C -- Yes --> D --> E --> F --> G
  C -- No  --> H

  %% -------- Lane styles --------
  style PA fill:#FFF4E5,stroke:#F6A609,stroke-width:1px
  style XL fill:#E8F5E9,stroke:#2E7D32,stroke-width:1px
  style FX fill:#F3E5F5,stroke:#7B1FA2,stroke-width:1px
  style OL fill:#FCE4EC,stroke:#C2185B,stroke-width:1px
