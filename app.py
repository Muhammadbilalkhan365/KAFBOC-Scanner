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
    
    # Blocklist for Headings, Cities, and Titles
    illegal_words = {
        'education', 'experience', 'summary', 'profile', 'skills', 'contact',
        'karachi', 'pakistan', 'lahore', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'university', 'college', 
        'certified', 'associate', 'manager', 'accountant', 'linkedin', 'email', 
        'phone', 'mobile', 'curriculum', 'resume', 'page', 'objective', 'hobbies', 
        'projects', 'mehmoodabad', 'clifton', 'gulshan', 'north', 'office', 'house',
        'no.', 'flat', 'street', 'road', 'sector', 'block', 'competencies', 'bookkeeper'
    }

    # Shuru ki lines uthayen
    lines = [l.strip() for l in text.split('\n') if l.strip()][:20]
    
    for line in lines:
        # Clean the line
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # Validation Logic:
        # A. Numbers bilkul na hon
        if any(char.isdigit() for char in cleaned): continue
        
        # B. URL ya Email na ho
        if any(x in low_line for x in ['http', 'www', '@', '.com', '.pk']): continue
        
        # C. Illegal words (Headings/Cities) na hon
        if any(word in low_line.split() for word in illegal_words): continue
        
        # D. Name Pattern (2 to 4 words, not too long)
        words = cleaned.split()
        if 2 <= len(words) <= 4 and 3 <= len(cleaned) <= 35:
            # E. Extra filter: Agar line "Contact" ya "Name:" se shuru ho rahi hai to ignore
            if not low_line.startswith(('name:', 'contact:', 'residence:')):
                return cleaned.title()
                        
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
        
        # Clean formatting
        text = text.replace('\t', ' ')
        
        # Email Extraction
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        
        return {
            "File Name": f_name,
            "Name": master_clean_name(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Processing Error", "Email": "N/A"}

# --- UI ---
files = st.file_uploader("Upload Files", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is filtering names...'):
        data = [extract_data(f) for f in files]
        df = pd.DataFrame(data)
    
    st.success("Cleaned Result:")
    st.table(df) # Table view zada clear hoti hai
    
    # Excel Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='CleanData')
        workbook = writer.book
        worksheet = writer.sheets['CleanData']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, col in enumerate(df.columns):
            worksheet.write(0, i, col, header_fmt)
        worksheet.set_column('A:C', 40)

    st.download_button("📥 Download Final Clean Excel", output.getvalue(), "KAFBOC_Master_Data.xlsx")
