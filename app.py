import streamlit as st
import fitz  # PyMuPDF
import docx
import re
import pandas as pd
import spacy
from io import BytesIO

# NLP Model Load karna
try:
    nlp = spacy.load("en_core_web_sm")
except:
    st.error("NLP Model load nahi ho saka. Requirements file check karein.")

st.set_page_config(page_title="KAFBOC Smart Miner", layout="wide")
st.title("📂 KAFBOC AI Data Extractor (v3.0)")

def advanced_clean_name(text):
    # 1. Ghair zaroori spaces theek karna (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # 2. NLP se Name dhoondna
    doc = nlp(text[:1000]) # Shuru ke 1000 characters kafi hain
    names = [ent.text for ent in doc.ents if ent.label_ == "PERSON"]
    
    # Block list (Jo cheezain name nahi ho saktin)
    block_list = ['karachi', 'pakistan', 'resume', 'curriculum', 'education', 'skills', 'experience', 'contact']
    
    for n in names:
        n_clean = " ".join(n.split())
        # Filter: Agar name boht chota/bada ho ya block list mein ho to skip karein
        if 3 < len(n_clean) < 30 and not any(b in n_clean.lower() for b in block_list):
            if not any(char.isdigit() for char in n_clean): # Phone number filter
                return n_clean.title()
                
    # Fallback: Agar NLP fail ho jaye to purana saf sutra logic
    lines = [l.strip() for l in text.split('\n') if l.strip()][:10]
    for line in lines:
        cleaned = " ".join(line.split())
        if 3 < len(cleaned) < 30 and not any(b in cleaned.lower() for b in block_list):
            if not any(char.isdigit() for char in cleaned):
                return cleaned.title()
                
    return "Not Found"

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
        
        # Robust Email Pattern
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        emails = re.findall(email_pattern, text)
        
        return {
            "File Name": file_name,
            "Name": advanced_clean_name(text),
            "Email": emails[0] if emails else "Not Found"
        }
    except:
        return {"File Name": file_name, "Name": "Error", "Email": "Failed"}

# --- UI ---
uploaded_files = st.file_uploader("Upload PDF/Word Files", accept_multiple_files=True)

if uploaded_files:
    results = [extract_info(f) for f in uploaded_files]
    df = pd.DataFrame(results)
    
    st.subheader("📋 Final Clean Data")
    st.table(df) # Table view zada saaf hoti hai
    
    # Excel Export
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
    
    st.download_button(
        label="📥 Download Professional Excel",
        data=output.getvalue(),
        file_name="KAFBOC_Final_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
