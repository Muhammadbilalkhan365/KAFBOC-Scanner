import streamlit as st
import fitz  # PyMuPDF
import docx
import re
import pandas as pd
from io import BytesIO

# Page Configuration
st.set_page_config(page_title="KAFBOC Data Miner", layout="wide")

st.title("📂 KAFBOC Professional Data Extractor")
st.info("System optimized for Excel (.xlsx) output.")

def clean_spaces(text):
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text) 
    return " ".join(text.split())

def is_valid_name(line):
    block_list = [
        'contact', 'education', 'experience', 'summary', 'profile', 'address', 
        'phone', 'mobile', 'resume', 'cv', 'competencies', 'skills', 'about', 
        'karachi', 'pakistan', 'linkedin', 'page', 'university', 'accountant'
    ]
    line_lower = line.lower()
    if any(word in line_lower for word in block_list): return False
    if re.search(r'[0-9]{5,}', line): return False 
    if len(line) < 3 or len(line) > 35: return False
    return True

def extract_info(uploaded_file):
    text = ""
    file_name = uploaded_file.name
    try:
        if file_name.endswith('.pdf'):
            bytes_data = uploaded_file.read()
            doc = fitz.open(stream=bytes_data, filetype="pdf")
            text = "".join([page.get_text() for page in doc])
        elif file_name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
        
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        emails = re.findall(email_pattern, text)
        
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        found_name = "Not Found"
        for line in lines:
            cleaned = clean_spaces(line)
            if is_valid_name(cleaned):
                found_name = cleaned.title()
                break
        
        return {"File Name": file_name, "Name": found_name, "Email": emails[0] if emails else "Not Found"}
    except Exception:
        return {"File Name": file_name, "Name": "Error", "Email": "Failed"}

# --- UI Interface ---
uploaded_files = st.file_uploader("Upload Resumes (PDF/Word)", accept_multiple_files=True)

if uploaded_files:
    results = [extract_info(f) for f in uploaded_files]
    df = pd.DataFrame(results)
    
    st.subheader("📋 Extracted Data Table")
    st.dataframe(df, use_container_width=True)
    
    # --- EXCEL DOWNLOAD LOGIC ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        # Formatting (Alignment theek karne ke liye)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format_left = workbook.add_format({'align': 'left'})
        worksheet.set_column('A:C', 30, format_left)
    
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Download as Excel (.xlsx)",
        data=processed_data,
        file_name="KAFBOC_Data_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.divider()
st.caption("Developed by Muhammad Bilal | KAFBOC Tech")
