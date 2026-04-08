import pandas as pd
import os

# --- SETTINGS (यहाँ आप अपनी फाइलों के नाम और कॉलम के नाम सेट करें) ---

# 1. फाइलों के नाम (File Names)
MAIN_FILE = 'Main_List.xlsx'          # आपकी मेन फाइल
PLACED_FILE = '2026 Placed Data.xlsx' # जिसमें Placed/Unplaced शीट्स हैं
SHEET_NAME = 'Placed'                 # Placed वाली शीट का नाम

# 2. कॉलम के नाम (Column Names)
# आपकी Main Sheet में जो नाम हैं
MAIN_EMAIL_COL = 'Email ID'      # मेन शीट में ईमेल वाले कॉलम का नाम बदलें अगर अलग हो
MAIN_PHONE_COL = 'Phone Number'  # मेन शीट में फोन वाले कॉलम का नाम बदलें अगर अलग हो

# आपकी Placed Sheet में जो नाम हैं
PLACED_EMAIL_COL = 'Email ID'    # प्लेस्ड शीट में ईमेल कॉलम का नाम
PLACED_PHONE_COL = 'Phone Number'# प्लेस्ड शीट में फोन कॉलम का नाम
PLACED_COMPANY_COL = 'Company'   # प्लेस्ड शीट में कंपनी नाम वाला कॉलम

def run_script():
    print("--- Process Started ---")
    
    # 1. फाइल लोड करना
    try:
        print(f"Loading {MAIN_FILE}...")
        df_main = pd.read_excel(MAIN_FILE)
        
        print(f"Loading {PLACED_FILE} (Sheet: {SHEET_NAME})...")
        df_placed = pd.read_excel(PLACED_FILE, sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"\nError: फाइल नहीं मिली या नाम गलत है।\nDetails: {e}")
        return

    # 2. डेटा सफाई (Data Cleaning) - ताकि स्पेस या फॉर्मेट की गलती न हो
    print("Cleaning Data for comparison...")
    
    # स्ट्रिंग में कन्वर्ट करें और स्पेस हटाएं
    df_main['clean_email'] = df_main[MAIN_EMAIL_COL].astype(str).str.strip().str.lower()
    df_main['clean_phone'] = df_main[MAIN_PHONE_COL].astype(str).str.strip() # .replace('.0', '') अगर नंबर में .0 आ रहा हो
    
    df_placed['clean_email'] = df_placed[PLACED_EMAIL_COL].astype(str).str.strip().str.lower()
    df_placed['clean_phone'] = df_placed[PLACED_PHONE_COL].astype(str).str.strip()
    
    # Placed डेटा से सिर्फ जरुरी चीज़ें रखेंगे (Mapping Dictionary बनाना तेज होता है)
    # हम दो डिक्शनरी बनाएंगे: एक Email से कंपनी खोजने के लिए, एक Phone से
    email_to_company = dict(zip(df_placed['clean_email'], df_placed[PLACED_COMPANY_COL]))
    phone_to_company = dict(zip(df_placed['clean_phone'], df_placed[PLACED_COMPANY_COL]))

    print("Comparing Data (Email OR Phone)...")

    # 3. मैच करना (Logic: पहले Email चेक करो, अगर न मिले तो Phone चेक करो)
    
    found_company_list = []
    status_list = []
    match_source_list = []

    for index, row in df_main.iterrows():
        m_email = row['clean_email']
        m_phone = row['clean_phone']
        
        company = None
        match_type = ""

        # Step A: Email से चेक करें
        if m_email in email_to_company:
            company = email_to_company[m_email]
            match_type = "Matched by Email"
        
        # Step B: अगर Email से नहीं मिला, तो Phone से चेक करें
        elif m_phone in phone_to_company:
            company = phone_to_company[m_phone]
            match_type = "Matched by Phone"
        
        # रिजल्ट लिस्ट में डालना
        if company:
            found_company_list.append(company)
            status_list.append("Placed")
            match_source_list.append(match_type)
        else:
            found_company_list.append("") # कुछ नहीं मिला
            status_list.append("Unplaced")
            match_source_list.append("")

    # 4. डेटा वापस Main Sheet में जोड़ना
    df_main['Is_Placed'] = status_list
    df_main['Placed_Company'] = found_company_list
    df_main['Match_Source'] = match_source_list

    # सफाई वाले अस्थायी कॉलम हटा दें
    df_main.drop(columns=['clean_email', 'clean_phone'], inplace=True)

    # 5. नई फाइल सेव करना
    output_filename = 'MWIDM Data.xlsx'
    df_main.to_excel(output_filename, index=False)
    
    print("\n--- Success! ---")
    print(f"नई फाइल '{output_filename}' बन गयी है।")
    print("अब आप इस फाइल को खोलें, 'Placed_Company' कॉलम देखें और मैन्युअली डिलीट करें।")

    # विंडोज को तुरंत बंद होने से रोकने के लिए
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    run_script()