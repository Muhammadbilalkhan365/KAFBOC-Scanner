import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Precision Miner", layout="wide")
st.title("📂 KAFBOC Professional Data Extractor (Precision Build)")

def is_valid_human_name(line):
    # A. Numbers bilkul nahi hone chahiye
    if any(char.isdigit() for char in line): return False
    
    # B. Ye woh keywords hain jo aapki screenshots mein ghalti kar rahe hain
    # Inhe hum sakhti se block karenge
    strict_blocklist = [
        'education', 'experience', 'summary', 'profile', 'skills', 'contact',
        'karachi', 'pakistan', 'lahore', 'address', 'about', 'communications', 
        'closing', 'reporting', 'modeling', 'accounting', 'certified', 'associate', 
        'manager', 'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 
        'cv', 'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 'expert',
        'key achievements', 'corporate tax', 'cma', 'process', 'management', 'accounts',
        'receivable', 'strong', 'communication', 'school', 'tabanis', 'accountancy',
        'having', 'international', 'serving', 'focused', 'professional'
    ]
    
    line_lower = line.lower()
    if any(bad_word in line_lower for bad_word in strict_blocklist):
        return False
        
    # C. Words count (Insaani naam aksar 2 se 3 alfaz ka hota hai)
    words = line.split()
    if len(words) < 2 or len(words) > 4:
        return False
        
    # D. Length (Bohat lamba sentence naam nahi ho sakta)
    if len(line) > 30:
        return False

    return True

def extract_precision_name(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Shuru ki 20 lines scan karein
    lines = [l.strip() for l in text.split('\n') if l.strip()][:20]
    
    for line in lines:
        cleaned = " ".join(line.split())
        # Agar saari checks pass ho jayein, to yehi Name hai
        if is_valid_human_name(cleaned):
            return cleaned.title()
                        
    return "Check Document"

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
        
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        return {
            "File Name": f_name,
            "Name": extract_precision_name(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Interface ---
files = st.file_uploader("Upload Resumes", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC System is refining data...'):
        data = [process_file(f) for f in files]
        df = pd.DataFrame(data)
    
    st.subheader("📋 Final Cleaned Report")
    st.dataframe(df, use_container_width=True)
    
    # Excel Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, col in enumerate(df.columns):
            worksheet.write(0, i, col, header_fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Excel Report", output.getvalue(), "KAFBOC_Final.xlsx")
