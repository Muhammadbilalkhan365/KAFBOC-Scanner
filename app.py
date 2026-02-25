import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Precision Miner", layout="wide")
st.title("📂 KAFBOC Professional Resume Parser")

def get_real_name(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Ye woh keywords hain jo aapki screenshots mein ghalti kar rahe hain
    exclusion_list = [
        'karachi', 'pakistan', 'lahore', 'education', 'skills', 'experience', 
        'summary', 'profile', 'contact', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'curriculum', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 
        'clifton', 'gulshan', 'street', 'house', 'flat', 'no.', 'sector', 'block',
        'competencies', 'bookkeeper', 'qualified', 'expert', 'remote', 'office'
    ]

    # Shuru ki 15-20 lines scan karein
    lines = [l.strip() for l in text.split('\n') if l.strip()][:20]
    
    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # Validation Rules:
        # A. Numbers, @, ya http ho to reject
        if any(char.isdigit() for char in cleaned) or "@" in low_line or "http" in low_line:
            continue
            
        # B. Words Count (Names usually have 2 to 3 words)
        words = cleaned.split()
        if not (2 <= len(words) <= 4):
            continue
            
        # C. Keyword Rejection (Sabse zaroori step)
        if any(bad_word in low_line for bad_word in exclusion_list):
            continue
            
        # D. Sentence Rejection (Agar lamba sentence hai to wo naam nahi hai)
        if len(cleaned) > 35:
            continue
            
        # E. Special Character Rejection (e.g. bullets ya symbols)
        if re.search(r'[•\*\-\|\/]', cleaned):
            continue

        # Agar saari checks pass ho jayein, to yehi Name hai
        return cleaned.title()
                        
    return "Check Manually"

def process_file(uploaded_file):
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
        email_matches = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        
        return {
            "File Name": f_name,
            "Name": get_real_name(text),
            "Email": email_matches[0] if email_matches else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Interface ---
files = st.file_uploader("Upload Resumes (Multiple Selection)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is filtering documents...'):
        data = [process_file(f) for f in files]
        df = pd.DataFrame(data)
    
    st.subheader("📋 Final Extracted Result")
    st.dataframe(df, use_container_width=True)
    
    # Excel Download formatting
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, val in enumerate(df.columns):
            worksheet.write(0, i, val, header_fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Excel File", output.getvalue(), "KAFBOC_Final_Report.xlsx")

st.divider()
st.caption("KAFBOC Tech Services - Accuracy First")
