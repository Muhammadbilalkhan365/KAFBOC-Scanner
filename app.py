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
    """CV se sirf candidate ka naam nikalne ke liye advanced logic"""
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Woh keywords jo Name column ko kharab kar sakte hain
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

    # CV ki shuruati lines mein naam dhoondna (Top 25 lines)
    lines = [l.strip() for l in text.split('\n') if l.strip()][:25]
    
    for line in lines:
        cleaned = " ".join(line.split())
        low_line = cleaned.lower()
        
        # A. Numbers bilkul nahi hone chahiye (Phone/Date filter)
        if any(char.isdigit() for char in cleaned): continue
        
        # B. Blocklist check (Titles/Headings filter)
        if any(bad_word in low_line for bad_word in blocklist): continue
        
        # C. Words count (Insaani naam aksar 2 se 4 alfaz ka hota hai)
        words = cleaned.split()
        if len(words) < 2 or len(words) > 4: continue
        
        # D. Length check (Bohat lamba sentence naam nahi ho sakta)
        if len(cleaned) > 35: continue

        # Agar saari checks pass ho jayein, to yehi Name hai
        return cleaned.title()
                        
    return "Check Document"

def process_resume(uploaded_file):
    """PDF aur Word files se text nikal kar Name/Email extract karna"""
    text = ""
    f_name = uploaded_file.name
    try:
        # File type check
        if f_name.endswith('.pdf'):
            file_bytes = uploaded_file.read()
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                text = "".join([page.get_text() for page in doc])
        elif f_name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
        
        # Email Extraction (Regex)
        email_list = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        email = email_list[0] if email_list else "Not Found"
        
        # Name Extraction (Using our validator)
        name = strict_name_validator(text)
        
        return {
            "File Name": f_name,
            "Name": name,
            "Email": email
        }
    except Exception as e:
        return {"File Name": f_name, "Name": "Error", "Email": str(e)[:15]}

# --- UI Interface ---
files = st.file_uploader("Upload Resumes (PDF/Word)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC AI is refining your data...'):
        results = []
        for f in files:
            results.append(process_resume(f))
        
        df = pd.DataFrame(results)
    
    st.success(f"Successfully processed {len(files)} files!")
    st.table(df)
    
    # Excel formatting for Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='CandidateData')
        
        workbook = writer.book
        worksheet = writer.sheets['CandidateData']
        
        # Professional Header Styling
        header_fmt = workbook.add_format({
            'bold': True, 
            'bg_color': '#1F4E78', 
            'font_color': 'white', 
            'border': 1,
            'align': 'center'
        })
        
        for i, col in enumerate(df.columns):
            worksheet.write(0, i, col, header_fmt)
            worksheet.set_column(i, i, 40) # Set column width

    st.download_button(
        label="📥 Download Professional Excel",
        data=output.getvalue(),
        file_name="KAFBOC_Extracted_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
