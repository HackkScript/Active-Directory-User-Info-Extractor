# 🧾Active-Directory-User-Info-Extractor
This Python script automates the process of retrieving detailed user account information from an Active Directory (AD) domain using the net user /domain command. It processes a list of usernames, gathers relevant information (such as full name, password set/expiry times, last logon, etc.), and saves the results to an Excel spreadsheet.

**🔧 Features**
`📋 Batch Processing`: Handles users in configurable batches (default: 500 users per batch).

`⚙️ Multithreading`: Uses concurrent threads to speed up data collection.

`📝 Excel Export`: Saves output in a structured .xlsx format using openpyxl.

`🔄 Resume Capability`: Automatically resumes from the last processed user if interrupted.

`🚫 Timeout and Retry`: Handles timeouts and logs failures for review.

`📊 Progress Bar`: Displays real-time progress using tqdm.

**📁 Input**
Provide a text file (users.txt) with one username per line:
  - user1
  - user2
  - user3

**📤 Output**
Generates an Excel file (e.g., user_details.xlsx) with columns like:

Username

Full Name

Account Active

Password Last Set (Date/Time)

Password Expires (Date/Time)

Password Required

Last Logon

Logon Script

Comment

**🚀 Usage**
  >python script.py users.txt [output_file.xlsx]

`users.txt`: Required input file containing usernames.

`output_file.xlsx`: Optional output file name (default: user_details.xlsx).

**📦 Requirements**

Python 3.6+

openpyxl

tqdm

**🛑 Notes**

This script must be run in a Windows environment connected to the appropriate domain.

Requires access rights to run net user /domain.

**📄 Logs**

`error_log.txt`: Logs failed queries or timeouts.

`resume_point.txt`: Stores the last processed user index for resume functionality.
