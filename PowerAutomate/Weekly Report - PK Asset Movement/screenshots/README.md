# Weekly Report – PK Asset Movement (Power Automate Flow)

This Power Automate flow runs weekly to prepare the PK Asset Movement report by copying the master file, running an Office Script to keep only the PK sheet, retrieving the updated file, and emailing it to stakeholders.

---

## 📌 Overview

- Schedule: Every Wednesday at 16:30 (UTC+07:00, Bangkok)
- Systems used: Power Automate, SharePoint, Excel Online (Office Scripts), Outlook
- Purpose:
  - Automate weekly PK asset movement report
  - Ensure consistent and clean data using an Office Script
  - Send the processed file automatically via email

---

## 🧩 Flowchart (Mermaid)

```mermaid
flowchart TD
    A[Recurrence Trigger\nWednesday 16:30\nUTC+07 Bangkok] --> B
    B[Copy file\nSource: Master Log 14NOV22.xlsx\nDestination: Archived folder\nIf exists: Replace] --> C
    C[Run Office Script\nScript name: KeepOnlyPKSheet\nFile ID: from Copy file] --> D
    D[Get file content\nFile ID: from Copy file] --> E
    E[Send email\nSubject: Weekly Report PK Asset Movement\nAttachment: processed file] --> F[End]
``
