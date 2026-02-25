import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Precision Miner", layout="wide")
st.title("📂 KAFBOC Professional Resume Parser")

def clean_and_extract_name(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # 2. Blocklist (In keywords ko Name column mein kabhi nahi aana chahiye)
    blocklist = [
        'karachi', 'pakistan', 'lahore', 'education', 'skills', 'experience', 
        'summary', 'profile', 'contact', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'cv', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 'expert',
        'having', 'international', 'focused', 'professional', 'senior', 'bookkeeper'
    ]

    # Shuru ki 15-20 lines scan karein
    lines = [l.strip() for l in text.split('\n') if l.strip()][:20]
    
    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # Validation Checks:
        # A. Numbers, @, ya Bullets ho to reject
        if any(char.isdigit() for char in cleaned) or "@" in low_line or "•" in cleaned:
            continue
            
        # B. Words Count (Names usually have 2 to 4 words)
        words = cleaned.split()
        if not (2 <= len(words) <= 4):
            continue
            
        # C. Keyword Rejection
        if any(bad in low_line for bad in blocklist):
            continue
            
        # D. Sentence Rejection (Agar boht lamba hai to wo name nahi)
        if len(cleaned) > 35:
            continue

        # Agar saari checks pass ho jayein, to yehi Name hai
        return cleaned.title()
                        
    return "Not Found"

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
            "Name": clean_and_extract_name(text),
            "Email": email_matches[0] if email_matches else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Interface ---
files = st.file_uploader("Upload Resumes", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is cleaning data...'):
        data = [process_file(f) for f in files]
        df = pd.DataFrame(data)
    
    st.subheader("📋 Extraction Result")
    # Table alignment for clear view
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

    st.download_button("📥 Download Excel Report", output.getvalue(), "KAFBOC_Final.xlsx")
