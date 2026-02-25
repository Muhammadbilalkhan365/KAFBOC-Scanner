import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Ultra-Miner", layout="wide")
st.title("📂 KAFBOC AI Data Extractor (Master Build)")

def master_clean_name(text):
    # 1. Spacing Fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Woh keywords jo Name column ko kharab kar rahe hain (Based on your screenshots)
    blacklist = [
        'karachi', 'pakistan', 'lahore', 'education', 'skills', 'experience', 
        'summary', 'profile', 'contact', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'curriculum', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 
        'clifton', 'gulshan', 'street', 'house', 'flat', 'no.', 'sector', 'block',
        'competencies', 'bookkeeper', 'qualified', 'expert', 'remote', 'office',
        'financial', 'international', 'markets', 'serving', 'focused', 'association'
    ]

    # Shuru ki 15-20 lines uthayen
    lines = [l.strip() for l in text.split('\n') if l.strip()][:20]
    
    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # Validation Rules:
        # A. Numbers bilkul na hon (Jo phone numbers aa rahe thay wo yahan se ruk jayenge)
        if any(char.isdigit() for char in cleaned): continue
        
        # B. URL, Email, ya Bullets (•) na hon
        if any(x in low_line for x in ['http', '@', '.com', '•', '|', '/']): continue
        
        # C. Blacklist check (Communications, Reporting wagera yahan se rukenge)
        if any(bad in low_line for bad in blacklist): continue
        
        # D. Name Pattern (2 to 4 words, length 3 to 35)
        words = cleaned.split()
        if 2 <= len(words) <= 4 and 3 <= len(cleaned) <= 35:
            return cleaned.title() # Format: Muhammad Bilal
                        
    return "Check Document"

def extract_data(uploaded_file):
    text = ""
    f_name = uploaded_file.name
    try:
        if f_name.endswith('.pdf'):
            with fitz.open(stream=uploaded_file.read(), filetype="pdf") as doc:
                text = "".join([page.get_text() for page in doc])
        elif f_name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
        
        # Email Extraction
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        
        return {
            "File Name": f_name,
            "Name": master_clean_name(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Layout ---
files = st.file_uploader("Upload Resumes", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is cleaning data...'):
        results = [extract_data(f) for f in files]
        df = pd.DataFrame(results)
    
    st.subheader("📋 Extraction Result")
    st.dataframe(df, use_container_width=True) # Saaf suthra table view
    
    # Excel Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, col in enumerate(df.columns):
            worksheet.write(0, i, col, header_fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Final Excel", output.getvalue(), "KAFBOC_Final.xlsx")
