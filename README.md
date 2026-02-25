# Daily-Data-Tracking-Email-Automation-for-OUTLOOK

## 📌 Overview

##This project is an automated data tracking pipeline that reads daily report files from multiple event and movie data sources, generates a summary report, and sends an automated email notification with the results.

It is designed to monitor data availability, track unique records, and provide daily visibility into live events and movie database updates.

🚀 Features

✅ Reads multiple CSV and Excel report files
✅ Tracks unique items and total rows per platform
✅ Generates a consolidated summary report
✅ Converts summary to HTML table
✅ Automatically sends email via Outlook
✅ Handles missing files gracefully
✅ Dynamic date-based file handling

🗂️ Supported Platforms

The script currently processes data from:

BookMyShow Main

BookMyShow Time & Date

District Insider

Skillbox

Neta Events

LiveYourCity

Movie DOD (Box Office)

🛠️ Tech Stack

Python

Pandas

Win32com (Outlook Automation)

Datetime

📁 Project Structure
📦 data-tracking-automation
 ┣ 📜 main_script.py
 ┣ 📜 README.md
 ┗ 📂 report_files
⚙️ How It Works

1️⃣ The script calculates yesterday’s date dynamically
2️⃣ Reads report files from predefined paths
3️⃣ Extracts:

Number of unique records

Total rows

4️⃣ Combines results into a summary DataFrame
5️⃣ Converts summary into HTML format
6️⃣ Sends automated email notification with report

📧 Email Output

The email includes:

Report date

Platform-wise summary table

Unique item counts

Total row counts

▶️ How to Run
1️⃣ Clone the repository
git clone https://github.com/your-username/data-tracking-automation.git
cd data-tracking-automation
2️⃣ Install dependencies
pip install pandas pywin32
3️⃣ Run the script
python main_script.py
🔧 Configuration

Update file paths inside the script:

file_paths = {
    'Platform_Name': 'your/local/path'
}

Update email recipients:

mail.To = 'your_email@example.com'
mail.CC = 'cc_emails@example.com'
⚠️ Requirements

Windows OS (required for Outlook automation)

Microsoft Outlook installed and configured

Python 3.8+

📈 Use Case

This tool is useful for:

✔ Data monitoring
✔ Daily ETL validation
✔ Reporting automation
✔ Data pipeline health checks
✔ Operations reporting

🧠 Future Improvements

Add logging system

Config file (YAML/JSON)

Docker support

Scheduler integration (Airflow / Cron)

Dashboard integration

Cloud storage support

👤 Author

Vrushank Dhande
Data Science Professional
