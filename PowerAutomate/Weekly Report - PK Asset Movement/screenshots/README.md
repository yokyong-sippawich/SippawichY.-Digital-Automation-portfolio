# Weekly Report – PK Asset Movement (Power Automate Flow)

This flow runs weekly to copy the master Excel file into the archive, execute an Office Script to keep only the PK sheet, fetch the updated file, and email it to stakeholders as an attachment.

---

## 📌 Overview

- Schedule: **Every Wednesday at 16:30 (UTC+07:00, Bangkok)**
- Stack: **Power Automate**, **SharePoint**, **Excel Online (Office Scripts)**, **Outlook**
- Purpose:
  - Automate weekly PK asset movement report preparation
  - Ensure consistent data processing with Office Script
  - Deliver the processed report to stakeholders reliably

---

## 🧩 Flowchart (Mermaid)

```mermaid
flowchart TD
    A[Trigger: Recurrence\nEvery Wednesday 16:30\nUTC+07:00 Bangkok] --> B

    B[SharePoint: Copy file\nFrom: Master Log 14NOV22.xlsx\nTo: Asset Movement Request Archived\nIf exists: Replace] --> C

    C[Excel Online: Run script\nScript: KeepOnlyPKSheet\nFile: ID from Copy file] --> D

    D[SharePoint: Get file content\nFile Identifier: ID from Copy file\nInfer Content Type: Yes] --> E

    E[Outlook: Send an email (V2)\nSubject: Weekly Report - PK Asset Movement\nAttach: processed Excel file] --> F[End]
