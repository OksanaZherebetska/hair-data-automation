# 💇‍♀️ Hair Data Automation

This repository contains a Python script used for automating daily data updates for a haircare collaboration project. The script runs scheduled queries on Google BigQuery, saves outputs to CSV and Excel, and sends an HTML summary email with key metrics.

> 🧪 This version is anonymized for portfolio/demo purposes — actual project names, emails, and paths have been removed.

---

## 🧠 What It Does

- ✅ Queries marketing & search term data from BigQuery  
- ✅ Saves output to `.csv` and refreshes an Excel template with up-to-date data  
- ✅ Sends an automated summary email with charts/tables (as HTML)  
- ✅ Cleans up old Excel files and kills rogue Excel processes (because Excel is Excel)  

---

## 🛠️ Tech Stack

- **Python** (pandas, datetime, win32com, psutil)
- **Google BigQuery** (via service account)
- **Excel automation** with `win32com.client`
- **Email automation** (via internal helper module)

---

## 🗂️ File Structure

---

*Note: helper modules should be added or mocked if you plan to run this script yourself.*

---

## ▶️ How to Use

1. Add your `service_account_key.json` file
2. Adjust `PROJECT`, paths, and output folder in `main.py`
3. Run the script:

```bash
python main.py

---

🧼 Disclaimer
This code is a cleaned-up version of a production-grade workflow and is intended to showcase automation logic, not real data or business logic. All sensitive information has been removed or masked.

👋 Author
Made with ☕ and 😅 by Oksana (Niekta)
