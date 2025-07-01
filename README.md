# Structural Column Data Exporter for Revit

A C#/.NET add-in for Autodesk Revit that automates the extraction of all structural column data into a clean, formatted, and ready-to-use Excel spreadsheet.

---

### 🎥 Demo

![Structural Column Exporter Demo](demo.gif)

---

### 🎯 The Problem

Structural engineers and project managers frequently need to create schedules and reports of all columns in a project for analysis, cost estimation, or fabrication. Manually creating and updating these reports from Revit is time-consuming and can lead to errors.

### ✨ The Solution

This add-in provides a one-click solution to instantly generate a comprehensive Excel report of all structural columns in the Revit model. It extracts key parameters and formats them into a professional-grade spreadsheet, saving hours of manual work.

---

### 🚀 Features

*   **One-Click Export:** Instantly exports data for all structural columns in the project.
*   **Formatted Excel Output:** Generates a professional `.xlsx` file with bolded headers and auto-sized columns using the EPPlus library.
*   **Smart Data Handling:** Correctly extracts and formats different parameter types, including text, numbers, and element references (like Base/Top Levels).
*   **Unit Conversion:** Automatically converts Revit's internal units into the project's user-facing display units.

---

### 🛠️ Tech Stack

*   **Languages & Frameworks:** C#, .NET Framework 4.8, Windows Form
*   **APIs & Libraries:** Revit API, **CloseXML**
*   **Development Tools:** Visual Studio 2022, Git
