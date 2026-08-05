# 📊 HRIS FMLA & Leave Claim Automation Tool

An enterprise-grade Python & Streamlit web application designed for HRIS and Benefits operations teams. This tool automates the auditing of formal leave-of-absence requests against system attendance logs, generating standardized executive Excel verification reports and formatted email replies.

---

## 🌟 Key Features

* **Automated Data Parsing:** Extracts employee details, leave case numbers, and required date ranges directly from raw email request text using regular expressions (Regex).
* **Flexible Attendance Log Parsing:** Seamlessly handles both Pipe-Delimited (`|`) and Tab-Delimited text exports copied directly from enterprise SAP time evaluation modules (`Cumulated Time Evaluation Results`).
* **Boundary & Audit Gap Verification:** Automatically audits whether the employee logged active shifts on the exact start and end boundary dates of the requested leave period.
* **Executive Excel Dashboards (`.xlsx`):**
  * Formatted attendance logs with zebra striping and custom monetary/numeric styling via `openpyxl`.
  * Built-in Excel formulas (`=SUM()`, `=COUNT()`) to dynamically aggregate total actual hours worked and shift counts.
  * Executive **Metrics Card & Audit Summary** panel highlighting key KPIs and missing date exceptions.
* **Dynamic Analyst Signature & Session Management:**
  * Custom user configuration inputs for Analyst Name and Job Title.
  * Powered by `st.session_state` — saves analyst credentials during active browser sessions and automatically resets upon closing the browser tab for multi-user data privacy.

---

## 🚀 How It Works

1. **Copy Leave Request:** Copy the incoming leave request email text from your Benefits Analyst inbox.
2. **Export Attendance Log:** Run your HRIS/SAP time balance report (`/TOTAL_HOURS` layout) for the requested leave date range and copy the results to your clipboard.
3. **Paste & Process:** Paste both inputs together into the web application and click **Process Claim & Generate Report**.
4. **Download & Reply:** Download the formatted Excel verification report (`<EmployeeID> - <LastName>.xlsx`) and copy the pre-formatted email response.

---

## 🛠️ Tech Stack & Dependencies

* **Language:** Python 3.10+
* **Frontend Framework:** [Streamlit](https://streamlit.io/)
* **Data Processing:** `pandas`, `io`, `re`
* **Excel Engine:** `openpyxl`
