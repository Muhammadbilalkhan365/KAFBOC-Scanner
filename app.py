import streamlit as st
import fitz
import docx
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAFBOC Advanced Miner", layout="wide")
st.title("📂 KAFBOC AI Data Extractor (Optimized Build)")

def clean_and_validate_name(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # Blocklist: In keywords wali lines ko kabhi Name nahi mana jayega
    blocklist = [
        'education', 'experience', 'summary', 'profile', 'skills', 'contact',
        'karachi', 'pakistan', 'address', 'about', 'communications', 'closing', 
        'reporting', 'modeling', 'accounting', 'university', 'college', 'certified', 
        'associate', 'manager', 'accountant', 'linkedin', 'email', 'phone', 
        'mobile', 'curriculum', 'resume', 'page', 'objective', 'hobbies', 'projects'
    ]

    # Shuru ki 15 lines scan karein (Naam hamesha top par hota hai)
    lines = [l.strip() for l in text.split('\n') if l.strip()][:15]
    
    for line in lines:
        cleaned = " ".join(line.split()) # Extra middle spaces hatana
        low_line = cleaned.lower()
        
        # Validation Logic:
        # - Line mein digits na hon
        # - Line blocklist mein na ho
        # - 2 se 4 alfaz hon (First/Last Name pattern)
        # - Length 3 se 30 chars ke darmiyan ho
        if not any(char.isdigit() for char in cleaned):
            if not any(word == low_line or word in low_line.split() for word in blocklist):
                words = cleaned.split()
                if 2 <= len(words) <= 4 and 3 <= len(cleaned) <= 30:
                    # Agar link ya email jaisa kuch hai to reject karein
                    if not any(x in low_line for x in ['http', 'www', '@', '.com']):
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
        
        # Robust Email Extraction
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        
        return {
            "File Name": f_name,
            "Name": clean_and_validate_name(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Layout ---
files = st.file_uploader("Upload Resumes", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC AI is cleaning your data...'):
        results = [process_resume(f) for f in files]
        df = pd.DataFrame(results)
    
    st.subheader("📋 Final Cleaned Table")
    st.dataframe(df, use_container_width=True)
    
    # Excel formatting
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for col, val in enumerate(df.columns):
            worksheet.write(0, col, val, header_fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Professional Excel Report", output.getvalue(), "KAFBOC_Final_Data.xlsx")
