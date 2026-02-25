import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Advanced Miner", layout="wide")
st.title("📂 KAFBOC Professional Data Extractor (Master Build)")

def master_name_validator(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Woh keywords jo aapki screenshots mein ghalti kar rahe hain (Strict Blocklist)
    blocklist = [
        'karachi', 'pakistan', 'lahore', 'education', 'skills', 'experience', 
        'summary', 'profile', 'contact', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'cv', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 'expert',
        'having', 'international', 'focused', 'professional', 'senior', 'bookkeeper',
        'key achievements', 'corporate tax', 'cma', 'process', 'management', 'accounts',
        'receivable', 'strong', 'communication', 'school', 'tabanis', 'accountancy'
    ]

    # Shuru ki 15-20 lines scan karein
    lines = [l.strip() for l in text.split('\n') if l.strip()][:25]
    
    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # Validation Rules:
        # A. Agar line mein numbers hain (Phone No) to reject
        if any(char.isdigit() for char in cleaned): continue
        
        # B. Agar line blocklist mein hai to reject
        if any(bad in low_line for bad in blocklist): continue
        
        # C. Agar lamba sentence hai (More than 35 chars) to reject
        if len(cleaned) > 35: continue
        
        # D. Name Pattern (Resumes mein naam aksar 2 se 3 words ka hota hai)
        words = cleaned.split()
        if 2 <= len(words) <= 4:
            # Check for special characters like / , - etc
            if not re.search(r'[/\\|•\*\(\)]', cleaned):
                return cleaned.title()
                        
    return "Check Document"

def process_resume(uploaded_file):
    text = ""
    f_name = uploaded_file.name
    try:
        if f_name.endswith('.pdf'):
            with fitz.open(stream=uploaded_file.read(), filetype="pdf") as doc:
                text = "".join([page.get_text() for page in doc])
        elif f_name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
        
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        return {
            "File Name": f_name,
            "Name": master_name_validator(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Interface ---
files = st.file_uploader("Upload Resumes (PDF/Word)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is refining your data...'):
        data = [process_resume(f) for f in files]
        df = pd.DataFrame(data)
    
    st.subheader("📋 Final Cleaned Report")
    st.dataframe(df, use_container_width=True)
    
    # Excel Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, val in enumerate(df.columns):
            worksheet.write(0, i, val, header_fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Professional Excel Report", output.getvalue(), "KAFBOC_Final.xlsx")
