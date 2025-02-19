# 📧 Email Extractor from Outlook

**Email Extractor** is a Python script that connects to Microsoft Outlook and extracts detailed email information, including sender addresses, recipients (To, CC, BCC), subject, body, and received time. The extracted data is then saved as a structured JSON file for easy use and analysis.

---

## 🚀 Features

- Extracts email metadata: **Subject**, **Sender**, **Recipients**, and **Body**.
- Resolves Exchange addresses to SMTP format.
- Categorizes recipients into **To**, **CC**, and **BCC**.
- Saves output as a formatted JSON file.
- Skips non-email items and handles errors gracefully.

---

## 🛠️ Requirements

Ensure you have the following installed on your system:

- **Python 3.x**  
- **Microsoft Outlook** (installed and configured)  
- Required Python packages:
  - `pywin32` for Outlook COM interaction  

Install the required package using:

```bash
pip install pywin32





