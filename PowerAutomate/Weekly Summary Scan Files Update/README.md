# Weekly Summary – Scan Files Update (Power Automate Flow)

This flow generates a weekly email summarizing scanned files for preventive and corrective maintenance work orders. It uses an Office Script to extract data from an Excel file stored in SharePoint and sends a formatted summary email to stakeholders.

---

## Overview

- Schedule: Weekly on Wednesday (UTC+07:00 Bangkok)
- Systems used: Power Automate, SharePoint, Excel Online (Office Scripts), Outlook
- Purpose:
  - Automate weekly summary of scanned Excel files
  - Convert extracted rows to HTML table
  - Send formatted email to stakeholders
  - Send "no data" email if no rows found

---

## Flowchart (Mermaid)

```mermaid
flowchart TB
    A[Recurrence (Weekly)] --> B[Run script: ReadSheets_LastFriToThu]
    B --> C{Rows > 0?}
    C -- Yes --> D[Create HTML table]
    D --> E[Compose IncludedSheetsString]
    E --> F[Compose EmailBody (HTML)]
    F --> G[Send summary email]
    C -- No --> H[Send no-data email]
