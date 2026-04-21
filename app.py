import streamlit as st
import fitz  # PyMuPDF
import docx
import re
import pandas as pd
from io import BytesIO

# Page Configuration
st.set_page_config(page_title="KAFBOC Precision Miner", layout="wide")
st.title("📂 KAFBOC Professional Data Extractor (Master Build)")

def strict_name_validator(text):
    """CV se candidate ka sahi naam nikalne ka logic"""
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Woh keywords jo Name nahi ho sakte
    blocklist = [
        'karachi', 'pakistan', 'lahore', 'education', 'skills', 'experience', 
        'summary', 'profile', 'contact', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'cv', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 'expert',
        'key achievements', 'corporate tax', 'cma', 'process', 'management', 'accounts',
        'receivable', 'strong', 'communication', 'school', 'tabanis', 'accountancy',
        'having', 'international', 'serving', 'focused', 'professional', 'bookkeeper'
    ]

    # Sirf top 25 lines scan karein
    lines = [l.strip() for l in text.split('\n') if l.strip()][:25]
    
    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # Filters
        if any(char.isdigit() for char in cleaned): continue
        if any(bad_word in low_line for bad_word in blocklist): continue
        
        words = cleaned.split()
        if len(words) < 2 or len(words) > 4: continue
        if len(cleaned) > 35: continue

        return cleaned.title()
                        
    return "Check Document"

def process_resume(uploaded_file):
    """File se text nikal kar Name aur Email extract karne ka function"""
    text = ""
    f_name = uploaded_file.name
    try:
        # File bytes ko aik baar read karke save karlein taake pointer issue na ho
        file_content = uploaded_file.read()
        
        if f_name.endswith('.pdf'):
            # stream=file_content aur filetype="pdf" ka istemal
            with fitz.open(stream=file_content, filetype="pdf") as doc:
                text = "".join([page.get_text() for page in doc])
        elif f_name.endswith('.docx'):
            # BytesIO ke zariye Word file read karna
            doc = docx.Document(BytesIO(file_content))
            text = "\n".join([para.text for para in doc.paragraphs])
        
        # Email Extraction
        email_list = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        email = email_list[0] if email_list else "Not Found"
        
        # Name Extraction using validator
        name = strict_name_validator(text)
        
        return {
            "File Name": f_name,
            "Name": name,
            "Email": email
        }
    except Exception as e:
        return {"File Name": f_name, "Name": "Error", "Email": str(e)[:20]}

# --- User Interface ---
files = st.file_uploader("Upload Resumes (PDF/Word)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC AI is extracting data...'):
        results = []
        for f in files:
            results.append(process_resume(f))
        
        df = pd.DataFrame(results)
    
    st.success(f"Processing Complete: {len(files)} files scanned.")
    st.table(df) # Dashboard par table display
    
    # Excel Formatting for Professional Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='CandidateData')
        
        workbook = writer.book
        worksheet = writer.sheets['CandidateData']
        
        # Header style (Blue background, White text)
        header_fmt = workbook.add_format({
            'bold': True, 
            'bg_color': '#1F4E78', 
            'font_color': 'white', 
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        for i, col in enumerate(df.columns):
            worksheet.write(0, i, col, header_fmt)
            worksheet.set_column(i, i, 35) # Column width adjust

    st.download_button(
        label="📥 Download Professional Excel",
        data=output.getvalue(),
        file_name="KAFBOC_Extracted_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
