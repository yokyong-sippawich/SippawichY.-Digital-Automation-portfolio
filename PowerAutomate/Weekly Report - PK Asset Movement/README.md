# Weekly Report - PK Asset Movement (Power Automate Flow)

## 📌 Overview
This flow automatically retrieves asset movement data, processes it, and generates a weekly report for PK assets. The automation reduces manual compilation time and ensures consistent reporting.

## 🎯 Objectives
- Automate recurring weekly asset report processing  
- Reduce manual data preparation work  
- Improve accuracy and consistency in reporting  
- Streamline the asset movement tracking workflow  

## ⚙️ Workflow Summary
1. **Trigger:** Scheduled (Weekly)  
2. **Source Data:** Excel / SharePoint list  
3. **Processing:**  
   - Filter data  
   - Clean and transform values  
   - Format report  
4. **Output:**  
   - Email report to stakeholders  
   - Save report to SharePoint folder  

## 📸 Screenshots
> Screenshots of the flow steps can be found here:  
[`screenshots`

## 📦 Exported Package  
Download the exported flow package:  
[`exportedpackage.zip`

## 🗂 Files Included
- `screenshots/` – All flow step visuals  
- `exported-package.zip` – Flow package for importing  
- `README.md` – Documentation for understanding the flow  

---

If you want to import this flow:
1. Download the `.zip` file  
2. Go to Power Automate → Solutions → Import  
3. Select the downloaded file

# Weekly Report – PK Asset Movement (Power Automate Flow)

This flow automatically processes the weekly PK Asset Movement Excel file, runs an Office Script to clean/filter the sheet, retrieves the processed file from SharePoint, and sends it to stakeholders as an email attachment.

---

## 📌 Overview

This automation runs **every Wednesday at 16:30 (UTC+07:00 – Bangkok)**.  
It performs the following actions:

1. Copy the master Excel file into the "Asset Movement Request Archived" folder (Replace if exists)  
2. Run an Office Script named **KeepOnlyPKSheet** on the copied file  
3. Retrieve the updated file content  
4. Email the processed file to stakeholders with a formatted date in the email body  

---

## 🧩 Flowchart Diagram (Mermaid)

```mermaid
flowchart TD
    A[Trigger: Recurrence<br/>Every Wednesday 16:30<br/>(UTC+07:00 – Bangkok)] --> B

    B[SharePoint: Copy file<br/>
      From:<br/>/Shared Documents/.../Asset Movement Request - Master Log 14NOV22.xlsx<br/>
      To:<br/>/Shared Documents/.../Asset Movement Request Archived/<br/>
      If exists: Replace] --> C

    C[Excel Online (Business): Run script<br/>
      Script: <b>KeepOnlyPKSheet</b><br/>
      File: ID from Copy file (body/id)] --> D

    D[SharePoint: Get file content<br/>
      File Identifier: ID from previous step<br/>
      Infer Content Type: Yes] --> E

    E[Outlook: Send an email (V2)<br/>
      To: Stakeholders<br/>
      Subject: Weekly Report - PK Asset Movement<br/>
      Attachment: Processed Excel file<br/>
      Body: Includes formatted date] --> F[End]
``
