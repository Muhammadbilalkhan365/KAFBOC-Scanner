import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Advanced Miner", layout="wide")
st.title("📂 KAFBOC AI Data Extractor (Precision Build)")

def filter_and_get_best_name(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # 2. Keywords jinhe Name nahi hona chahiye (Strict List)
    bad_keywords = [
        'karachi', 'pakistan', 'lahore', 'education', 'skills', 'experience', 
        'summary', 'profile', 'contact', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'curriculum', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 
        'gulshan', 'north', 'office', 'house', 'no.', 'flat', 'street', 'road', 
        'sector', 'block', 'competencies', 'bookkeeper', 'qualified', 'expert'
    ]

    # Shuru ki 15-20 lines scan karein
    lines = [l.strip() for l in text.split('\n') if l.strip()][:20]
    
    best_candidate = "Check Document"
    max_score = -100

    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        score = 0
        
        # Validation Rules:
        # A. Numbers ya @/http ho to seedha reject
        if any(char.isdigit() for char in cleaned) or "@" in low_line or "http" in low_line:
            continue
            
        # B. Words Count (Names usually have 2 to 4 words)
        words = cleaned.split()
        if 2 <= len(words) <= 4:
            score += 50
        else:
            score -= 30
            
        # C. Keyword Penalty
        if any(word in low_line for word in bad_keywords):
            score -= 100
            
        # D. Length Check
        if 4 <= len(cleaned) <= 35:
            score += 20
        else:
            score -= 50

        # Agar is line ka score ab tak ka sabse behtar hai to save karein
        if score > max_score and score > 0:
            max_score = score
            best_candidate = cleaned.title()
                        
    return best_candidate

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
            "Name": filter_and_get_best_name(text),
            "Email": email_matches[0] if email_matches else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Interface ---
files = st.file_uploader("Upload Files (PDF/Word)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is filtering names...'):
        data = [process_file(f) for f in files]
        df = pd.DataFrame(data)
    
    st.subheader("📋 Extraction Result")
    st.dataframe(df, use_container_width=True) # Dataframe view is clean
    
    # Excel formatting
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, val in enumerate(df.columns):
            worksheet.write(0, i, val, fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Final Excel Report", output.getvalue(), "KAFBOC_Final_Report.xlsx")
