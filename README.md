# Google Sheets Data Aggregator & Auditor

A powerful automation tool designed to streamline data workflows by processing multiple Google Sheets within a directory. This project automates the tedious process of consolidating scattered spreadsheets into a single, comprehensive overview while simultaneously auditing the data to isolate conflicts and discrepancies.

---

## 🚀 Features

* **Folder-Wide Data Aggregation:** Automatically scans a designated Google Drive folder, extracts relevant data from multiple Google Sheets, and compiles them into a single, comprehensive master detail sheet.
* **Intelligent Discrepancy Auditing:** Cross-references data points across all source files to instantly flag conflicting or disputed entries, outputting them into a dedicated comparison sheet for quick resolution.
* **Automated Workflow:** Eliminates manual file downloading, copying, and pasting, saving hours of data reconciliation effort.

---

## 🛠️ How It Works

1. **Ingestion:** The script securely connects to your target folder containing the source sheets.
2. **Aggregation:** Data is parsed, cleaned, and merged into a `Master_Detail_Sheet`.
3. **Conflict Detection:** An audit engine compares overlapping data fields to find mismatches (e.g., differing values for the same record ID across different files).
4. **Output:** Generates or updates two distinct sheets:
   * **Detail Sheet:** The unified data repository.
   * **Comparison/Dispute Sheet:** A targeted list of rows requiring manual review or correction.

---

## ⚙️ Prerequisites & Tech Stack

* **Language:**  Google Apps Script
* **APIs Used:** Google Sheets API, Google Drive API

---

## 🚀 Getting Started

### 1. Installation
Clone the repository to your local machine:
```bash
git clone [https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git)
cd YOUR_REPOSITORY_NAME


