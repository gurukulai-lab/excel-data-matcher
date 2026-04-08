📊 Excel Data Matcher
A Python automation tool that compares two Excel sheets and identifies matching records based on email and phone number.

🚀 Features
📂 Compare two Excel files
🔍 Match data using Email OR Phone
🧹 Data cleaning (remove spaces, normalize format)
📊 Identify Placed and Unplaced records
🏢 Extract company details for matched entries
⚡ Fast and automated processing

🛠️ Tech Stack
Python
Pandas
Excel (openpyxl)

📂 Project Structure
Excel_Data_Matcher/
│── main.py
│── Main_List.xlsx
│── Placed_Data.xlsx
│── README.md

⚙️ How It Works
Load main dataset
Load placed students dataset
Clean and normalize data
Match records using Email or Phone
Generate new Excel file with results

▶️ Usage
Run the script:
python main.py

📌 Requirements
Install dependencies:
pip install pandas openpyxl

📊 Output
New Excel file generated
Columns added:
Is_Placed
Placed_Company
Match_Source

⚡ Use Cases
🎓 Student placement tracking
📊 Data comparison and cleaning
🏢 HR data processing
📈 Excel automation

⚠️ Disclaimer
This project is for educational and automation purposes only.

👨‍💻 Author
Built by GurukulAI Lab 🚀
