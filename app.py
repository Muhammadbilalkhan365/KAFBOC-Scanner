import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Precision Miner", layout="wide")
st.title("📂 KAFBOC Professional Resume Parser")

def get_best_name_candidate(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Blocklist: In alfaz ko Name column mein nahi aana chahiye
    blocklist = [
        'karachi', 'pakistan', 'lahore', 'education', 'skills', 'experience', 
        'summary', 'profile', 'contact', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'cv', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 'expert',
        'having', 'international', 'focused', 'professional', 'senior', 'bookkeeper'
    ]

    lines = [l.strip() for l in text.split('\n') if l.strip()][:25]
    best_name = "Not Found"
    highest_score = -100

    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        current_score = 0
        
        # --- SCORE CALCULATION ---
        # A. Numbers, @, ya special symbols hon to score boht kam kar do
        if any(char.isdigit() for char in cleaned) or "@" in low_line or "•" in cleaned:
            current_score -= 200
            
        # B. Words Count: Names usually have 2 or 3 words (Muhammad Bilal)
        words = cleaned.split()
        if 2 <= len(words) <= 3:
            current_score += 50
        elif len(words) == 1:
            current_score -= 20 # Single words like "Education" get penalty
            
        # C. Keyword Penalty: Agar blocklist ka lafz ho to seedha reject
        if any(bad in low_line for bad in blocklist):
            current_score -= 150
            
        # D. Case Sensitivity: All CAPS names are common in resumes
        if cleaned.isupper() and len(cleaned) > 5:
            current_score += 20
            
        # E. Length Check: Ideal name length 5 to 30 chars
        if 5 <= len(cleaned) <= 30:
            current_score += 30
        else:
            current_score -= 50

        # --- HIGHEST SCORE SELECTION ---
        if current_score > highest_score:
            highest_score = current_score
            best_name = cleaned.title()
                        
    return best_name if highest_score > 10 else "Check Document"

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
            "Name": get_best_name_candidate(text),
            "Email": email_matches[0] if email_matches else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Interface ---
files = st.file_uploader("Upload Resumes (Multiple Selection)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is refining data...'):
        data = [process_file(f) for f in files]
        df = pd.DataFrame(data)
    
    st.subheader("📋 Extraction Result")
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
