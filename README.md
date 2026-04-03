# UOP-Science-GPA-Analyzer-Predictive-dashboard
A professional Excel-based GPA analyzer and predictive dashboard for Faculty of Science, University of Peradeniya, featuring weighted GPA calculator for BSc &amp; BSc Hons degrees and target grade forecasting.
# 📊 UOP Science Faculty GPA Analyzer & Predictive Dashboard

An advanced, automated Excel-based analytical tool designed specifically for students at the **Faculty of Science, University of Peradeniya (UOP)**. This project demonstrates high-level proficiency in **Excel Automation, VBA, and Predictive Data Modeling**.

## 🚀 Key Features

### **1. Intelligent Data Entry Dashboard**
* **Automated Course Recognition:** Input a course code (e.g., `MAT2013`), and the system automatically identifies:
* **Subject:** (e.g., `MAT` = Mathematics).
* **Academic Year:** Identified via the 1st digit.
* **Credit Value:** Identified via the last digit.
* **VBA Powered Submission:** A custom **"SUBMIT"** button developed using **Excel VBA Macros** to transfer data from the dashboard to the main database instantly.

### **2. Academic Weightage Logic**
Programmed with the official UOP Science Faculty weighting criteria for 100% precision:
* **BSc (General):** 1st Year (20%), 2nd Year (40%), 3rd Year (40%).
* **BSc Hons:** 1st Year (20%), 2nd Year (20%), 3rd Year (30%), 4th Year (30%).

### **3. Smart Target GPA Forecaster (Predictive Analysis)**
* **Goal Tracking:** Users can set a target GPA, and the tool calculates the required GPA for the **remaining academic years**.
* **Dynamic Simulation:** Allows "What-if" scenarios to see how a potential grade in a future semester affects the final degree classification (e.g., First Class, Second Upper).
* **One-Touch Reset:** A VBA-based **"RESET"** function to clear all simulations.

---

## 📐 Technical Architecture & Formulas

The core of this tool relies on advanced logical functions and error handling to ensure data integrity:

### **Advanced Logical Handling (Nested IFs)**
Used to determine degree eligibility and honors classification.
Example:
`=IF(D51=0, IF(A51=0, D49, (D49-(0.2*A51))/0.8), D51)`

### **Data Retrieval & Cleaning**
* **VLOOKUP & IFERROR:** Used to fetch grade points dynamically and suppress errors for a clean UI.
* **String Manipulation:** Uses `LEFT`, `MID`, and `RIGHT` functions to decode course codes.

### **Automation (VBA)**
The tool utilizes Macro-enabled automation for the Submit and Reset functionalities, providing a software-like experience within Excel.

---

## 🛠️ Tech Stack
* **Microsoft Excel:** Advanced Data Modeling.
* **VBA (Visual Basic for Applications):** Backend Automation.
* **Data Visualization:** Interactive Dashboards with Slicers and Conditional Formatting.

---

## 📂 How to Use
1. **Download** the `.xlsm` file.
2. **Enable Macros** to allow the Submit/Reset buttons to function.
3. **Enter Course Code & Result:** Select from dropdowns and click Submit.
4. **Analyze:** View your progress on the interactive dashboard.
