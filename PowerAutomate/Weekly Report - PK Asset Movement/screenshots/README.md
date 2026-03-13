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
flowchart TB
  %% --------------- LANES ---------------
  subgraph PA[Power Automate]
    direction TB
    A[⏰ Recurrence\nWed 16:30\nUTC+07]
  end

  subgraph SP[SharePoint]
    direction TB
    B[📁 Copy master to archive\nReplace if exists]
    D[📥 Get file content\nby file ID]
  end

  subgraph XL[Excel Online]
    direction TB
    C[🧮 Run script\nKeepOnlyPKSheet]
  end

  subgraph OL[Outlook]
    direction TB
    E[📧 Send email\nAttach processed file]
  end

  %% --------------- FLOW ---------------
  A --> B --> C --> D --> E

  %% --------------- LANE STYLES ---------------
  style PA fill:#FFF4E5,stroke:#F6A609,stroke-width:1px
  style SP fill:#E8F3FF,stroke:#1E88E5,stroke-width:1px
  style XL fill:#E8F5E9,stroke:#2E7D32,stroke-width:1px
  style OL fill:#FCE4EC,stroke:#C2185B,stroke-width:1px

  %% --------------- NODE CLASSES (optional) ---------------
  classDef control fill:#FFE8CC,stroke:#EB8A00,color:#333,stroke-width:1px
  classDef sharepoint fill:#DDEEFF,stroke:#1E88E5,color:#0D47A1,stroke-width:1px
  classDef excel fill:#DFF5E1,stroke:#2E7D32,color:#1B5E20,stroke-width:1px
  classDef outlook fill:#FFD9E2,stroke:#C2185B,color:#880E4F,stroke-width:1px

  class A control
  class B,D sharepoint
  class C excel
  class E outlook
