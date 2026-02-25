import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Precision Miner", layout="wide")
st.title("📂 KAFBOC Professional Data Extractor (Master Build)")

def strict_name_validator(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Woh keywords jo Name column ko kharab kar rahe hain (Strict Blocklist)
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

    # Shuru ki 15-20 lines scan karein (Naam hamesha top par hota hai)
    lines = [l.strip() for l in text.split('\n') if l.strip()][:25]
    
    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # A. Numbers bilkul nahi hone chahiye (Phone No filter)
        if any(char.isdigit() for char in cleaned): continue
        
        # B. Blocklist check (Headings aur Titles filter)
        if any(bad_word in low_line for bad_word in blocklist): continue
        
        # C. Words count (Insaani naam aksar 2 se 3 alfaz ka hota hai)
        words = cleaned.split()
        if len(words) < 2 or len(words) > 4: continue
        
        # D. Length check (Bohat lamba sentence naam nahi ho sakta)
        if len(cleaned) > 35: continue

        # Agar saari checks pass ho jayein, to yehi Name hai
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
        
        # Proper Email Extraction
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        
        return {
            "File Name": f_name,
            "Name": strict_name_cleaner(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Interface ---
files = st.file_uploader("Upload Files (PDF/Word)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC AI is refining your data...'):
        data = [process_resume(f) for f in files]
        df = pd.DataFrame(data)
    
    st.success("Clean Data Prepared!")
    st.table(df) # Proper Table alignment for clear view
    
    # Excel formatting for Professional Presentation
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, val in enumerate(df.columns):
            worksheet.write(0, i, val, header_fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Professional Excel", output.getvalue(), "KAFBOC_Final.xlsx")
