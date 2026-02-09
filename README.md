# BI in Excel: Automated Dashboards using Power Query and Power Pivot 

[![BI in Excel – Automated Dashboard Preview](Snapshots/Dashboard_Preview.png)](https://github.com/quzeem91/BI-in-Excel-Automated-Dashboard-using-Power-Query-and-Power-Pivot/releases/download/v1.0/Automated_Dashboard_File.xlsx)

**Click the image above to download and interact with the Dashboard in Excel.**



---
## Business Problem 🧩

Many **small businesses and lean teams** want **Power BI–style reporting**, automated data refresh, clean data models, and interactive dashboards, but face **budget constraints** that make Power BI licensing impractical.

In addition:

- Excel is still the **default reporting tool** in many organizations
- Teams want **familiar tools** without sacrificing BI best practices

This creates a gap 👉:   *How can teams build scalable, automated, and visually compelling BI reports using only Excel?*



---

## Solution 💡

This project demonstrates how **Microsoft Excel** can serve as a full-fledged BI tool, by leveraging:

- **Power Query** for data ingestion and transformation
- **Power Pivot** for data modeling and DAX calculations
- **Wireframing** to deliver a Power BI–like user experience

The result is a **fully automated Excel BI dashboard** that looks, feels, and behaves like Power BI without licensing costs.

---

## How to Build ⚙️ 

### Step 1: Connecting to Multiple Data Sources with Power Query

A core feature of any BI tool is the ability to connect to **multiple data sources**, both on-premise and cloud-based.

With **Power Query** now fully built in Excel, this enabled the project workflow to:

- Connect to **multiple datasets**
- Automate **data cleaning and transformation** processes
- Enable a **one-click refresh** of the dashboard

**Data Sources Used:**

- **Folder connection** containing transactional fact tables 
- **Multiple CSV files** serving as lookup/ dimension tables 

This mirrors how Power BI uses Power Query for ETL processes.

---

### Step 2: Data Modeling with Power Pivot

**Power Pivot** is Excel’s built-in data modeling engine and functions similarly to the Power BI semantic model.

In this step:

- Relationships were created between fact and lookup tables
- A **snowflake schema** was implemented
- **Calculated columns, measures, and a calendar table** were built using **DAX**
- All calculations were centralized in the model (not visuals)

>  *The above functionality was once exclusive to Power BI but is fully achievable in Excel via Power Pivot.*

**Model Snapshot:** 

![Data Model](Snapshots/Data_Model.png)

---

### Step 3: Wireframing for a Power BI Look & Feel

Wireframing was critical to achieving a **Power BI–style UX** within Excel.

This step involved:

- Drawing inspiration from modern Power BI dashboards
- Defining visual hierarchy and spacing
- Planning slicer placement and interaction flow
- Designing the dashboard layout in **PowerPoint**



This ensured the final Excel dashboard felt **intentional, clean, and professional**.
See Wireframe PPT and Inspo below

**Wireframe PPT**

- ![Wireframe](Snapshots/Wireframe_Snapshot.png)

**Wireframe Inspo**

- ![Wireframe-Inspo](Wireframe/Wireframe_Inspo.png)

---

### Step 4: Bringing It All Together

With the data model and measures in place:

- PivotTables were created directly from the **Power Pivot model**
- Charts and visuals were layered on top
- Slicers were connected across visuals
- The dashboard was assembled to match the wireframe

The result:

- Interactive visuals
- Centralized calculations
- Seamless refresh from source to dashboard

[![BI in Excel – Automated Dashboard Preview](Snapshots/Dashboard_Preview.png)](https://github.com/quzeem91/BI-in-Excel-Automated-Dashboard-using-Power-Query-and-Power-Pivot/releases/download/v1.0/Automated_Dashboard_File.xlsx)

**Click the image above to download and interact with the Dashboard in Excel.**

---

## Limitations ⚠️

- **No RLS**: Excel cannot enforce user-based data filtering or access roles
- **Model Maintenance**: Managing relationships, DAX, Dependencies, and Calculations is more tedious in Excel
- **No Enable Load Toggle**: Unused/ helper queries must be deleted, else they get loaded to the model as opposed to not loading them into the data model
- **Collaboration**: Collaboration is not as seamless as Power BI Service



Despite these, it remains a **strong alternative for small teams and budget-constrained environments**.

---

## Live Project Build 🎥

This project was built **from scratch in a live session**.

**Watch the full build webinar:▶️** 
*[here](https://www.youtube.com/watch?v=cZD9RHfbz04&t=22s)*

---

## Tools & Technologies 🛠️

- Microsoft Excel
- Power Query
- Power Pivot
- DAX
- PowerPoint (Wireframing)

---

## Key Takeaway 📌

> **Business Intelligence is a mindset, not a tool.**
> With the right workflow, Excel can deliver **BI-grade reporting** that scales far beyond traditional spreadsheets.
