# -*- coding: utf-8 -*-
"""
Created on Sun Apr 13 18:33:31 2025

@author: Oksana.Zherebetska
"""

# -*- coding: utf-8 -*-
"""
Automated script to:
- Query data from BigQuery
- Save outputs to CSV and Excel
- Send HTML summary email

Structure based on real production script. Cleaned for portfolio use.
"""

import sys
import os
import time
import psutil
import pandas as pd
from datetime import datetime, date, timedelta
import win32com.client

# Insert project path if needed
sys.path.insert(1, r'C:\your\path\to\toolbox')

# Custom modules (to be provided separately)
from bigquery_ops import BigQueryLoader
from send_email import send_email_no_attachment

from google.oauth2 import service_account

# === BigQuery setup ===
key_path = 'path/to/service_account_key.json'
PROJECT = 'your-project-id'
credentials = service_account.Credentials.from_service_account_file(key_path)
bq_loader = BigQueryLoader(PROJECT, credentials)

# === Prepare time variables ===
start = date.today() - timedelta(days=1)
year = start.year
last_year = year - 1

# === Output path ===
output_path = r'C:\your\project\output'

# === Kill any remaining Excel processes ===
def kill_excel_process():
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] == 'EXCEL.EXE':
            proc.terminate()

kill_excel_process()

# === SQL queries (anonymized) ===
query_marketing_data = """
SELECT *
FROM `your-project.dataset.marketing_data`
WHERE Year >= {}
""".format(last_year)

query_search_data = """
SELECT
    Search_Term_Type,
    Search_Term,
    Month,
    TY_Search_Vol,
    LY_Search_Vol,
    TY_LM_Search_Vol
FROM (
    SELECT
        Search_Term_Type,
        Search_Term,
        Month,
        SUM(TY_Search_Vol) AS TY_Search_Vol,
        SUM(LY_Search_Vol) AS LY_Search_Vol,
        SUM(TY_LM_Search_Vol) AS TY_LM_Search_Vol,
        ROW_NUMBER() OVER (
            PARTITION BY Search_Term_Type, Month
            ORDER BY SUM(TY_Search_Vol) DESC
        ) AS rn
    FROM `your-project.dataset.search_trends`
    WHERE EXTRACT(YEAR FROM Date) = {}
    GROUP BY ALL
)
WHERE rn <= 500
ORDER BY 1, 3, 4 DESC
""".format(year)

# === Run main logic inside a function ===
def SCRIPT():
    completion_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # === Load data from BigQuery ===
    df_marketing = bq_loader.exec_bq_query(query_marketing_data)
    df_search = bq_loader.exec_bq_query(query_search_data)

    # === Save to CSV ===
    df_marketing.to_csv(os.path.join(output_path, 'marketing_data.csv'), index=False)
    df_search.to_csv(os.path.join(output_path, 'search_data.csv'), index=False)
    print("CSV files saved.")

    # === Excel refresh ===
    excel_template = os.path.join(output_path, 'template.xlsx')
    new_excel_name = f"Data Report {completion_time}.xlsx"
    new_excel_path = os.path.join(output_path, new_excel_name)

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.Interactive = False

        workbook = excel.Workbooks.Open(os.path.abspath(excel_template))
        sheet = workbook.Sheets(1)
        sheet.Range("B4").Value = completion_time
        print("Excel opened, refreshing data...")
        workbook.RefreshAll()
        time.sleep(5)
        workbook.SaveAs(new_excel_path)
        workbook.Close()

        # Cleanup old Excel files (except template and current)
        for file in os.listdir(output_path):
            fpath = os.path.join(output_path, file)
            if file.endswith(".xlsx") and file not in [os.path.basename(excel_template), os.path.basename(new_excel_path)]:
                try:
                    os.remove(fpath)
                    print(f"Deleted old file: {file}")
                except Exception as e:
                    print(f"Failed to delete {file}: {e}")

        print(f"Excel report saved as {new_excel_path}")
    except Exception as e:
        print(f"Excel error: {e}")
    finally:
        kill_excel_process()
        if 'excel' in locals():
            excel.Quit()
            excel = None

    # === Summary queries for email output ===
    summary_query_1 = """
    SELECT Time_Period, Start_Date, End_Date, 
           CONCAT('£', FORMAT("%'.2f", SUM(Sales))) AS Sales
    FROM `your-project.dataset.marketing_data`
    WHERE Year = {}
    GROUP BY ALL
    ORDER BY End_Date DESC
    LIMIT 10
    """.format(year)

    summary_query_2 = """
    SELECT Date,
           SUM(TY_Search_Vol) AS TY_Search_Vol,
           SUM(LY_Search_Vol) AS LY_Search_Vol,
           SUM(TY_LM_Search_Vol) AS TY_LM_Search_Vol
    FROM `your-project.dataset.search_trends`
    WHERE Date BETWEEN DATE_SUB(CURRENT_DATE(), INTERVAL 10 DAY) AND DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY)
    GROUP BY ALL
    ORDER BY Date DESC
    """

    df_summary_1 = bq_loader.exec_bq_query(summary_query_1)
    df_summary_2 = bq_loader.exec_bq_query(summary_query_2)

    summary_html_1 = df_summary_1.to_html(index=False)
    summary_html_2 = df_summary_2.to_html(index=False)

    # === Send email summary ===
    send_list = ('your_email@example.com',)
    subject = f"Auto-email: Data update completed {completion_time}"
    body = f"""
    <html>
      <body>
        <p>Hi team,</p>
        <p>Data update completed at {completion_time}.</p>
        <p><strong>Marketing Summary:</strong><br>{summary_html_1}</p>
        <p><strong>Search Summary:</strong><br>{summary_html_2}</p>
        <p>Regards,<br>Data Bot 🤖</p>
      </body>
    </html>
    """

    send_email_no_attachment(send_list, subject, body)
    print("Email sent.")

# === Error handling wrapper ===
try:
    SCRIPT()
except Exception as err:
    subject = f"Auto-email: Data update FAILED {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    error_body = f"""
    <html>
      <body>
        <p>Script failed with error:</p>
        <pre>{err}</pre>
      </body>
    </html>
    """
    send_email_no_attachment(('your_email@example.com',), subject, error_body)
    print("Failure email sent.")
