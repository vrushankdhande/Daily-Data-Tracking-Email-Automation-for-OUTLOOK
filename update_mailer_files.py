# ==============================
# 📦 IMPORT LIBRARIES
# ==============================
import pandas as pd
import time
import win32com.client as win32
from datetime import datetime, timedelta

# ==============================
# 📅 DATE SETUP
# ==============================
today = datetime.today()
yesterday = today - timedelta(days=1)
today_str = today.strftime("%Y-%m-%d")
yesterday_str_file = yesterday.strftime("%Y_%m_%d")
yesterday_str_report = yesterday.strftime("%Y-%m-%d")
print("Today:", today_str)
print("Yesterday:", yesterday_str_file)

# ==============================
# 📂 FILE CONFIGURATION
# Format → extension + unique column
# ==============================
file_config = {
    'your_file_entension': ('.csv', 'URL'),
    'your_file_entension1': ('.csv', 'url')
}

# ==============================
# 📁 FILE PATHS
# ==============================
file_paths = {
    'your_file_entension': f'our_path_location_{yesterday_str_file}',
    'your_file_entension1': f'our_path_location__{yesterday_str_file}'}

# ==============================
# 📊 PROCESS FILES
# ==============================
summary_list = []

for platform, path in file_paths.items():
    try:
        print(f"Processing → {platform}")

        file_extension, unique_column = file_config[platform]
        full_path = path + file_extension

        # Read file based on extension
        if file_extension == '.csv':
            df = pd.read_csv(full_path)
        else:
            df = pd.read_excel(full_path)

        # Create summary row
        temp_df = pd.DataFrame([{
            'Platform': platform,
            'Unique_items': df[unique_column].nunique(),
            'Total_rows': len(df),
            'Report_date': yesterday_str_report
        }])

        summary_list.append(temp_df)

    except Exception as e:
        print(f"❌ Unable to process file for platform: {platform}")
        print("Error:", e)

# ==============================
# 📊 FINAL OUTPUT DATAFRAME
# ==============================
final_output_df = pd.concat(summary_list).reset_index(drop=True)

print("\nFinal Summary:")
print(final_output_df)

# ==============================
# 📨 CONVERT DATAFRAME TO HTML
# ==============================
df_html = final_output_df.to_html(index=False, border=1)

# ==============================
# 📧 CREATE OUTLOOK EMAIL
# ==============================
outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)

html_body = f"""
<html>
  <body>
    <p>Hi User,</p>
    <p>Below is the tracked data from Server Database for date: <b>{today_str}</b>:</p>
    <p>The data is for Live Events and Movie Database DOD.</p>
    {df_html}
    <p>Thank you,<br>User</p>
  </body>
</html>
"""

# ==============================
# 📬 EMAIL DETAILS
# ==============================
mail.To = 'your_email@example.com'
mail.CC = 'cc_emails@example.com'
mail.Subject = f'Tracking status for date {today_str}'
mail.HTMLBody = html_body

# ==============================
# 🚀 SEND EMAIL
# ==============================
time.sleep(3)
mail.Send()


print("✅ Mail sent successfully.")
