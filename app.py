import streamlit as st
import fitz
import docx
import re
import pandas as pd
import spacy
from io import BytesIO

# AI Model Load
@st.cache_resource
def load_nlp():
    try:
        return spacy.load("en_core_web_sm")
    except:
        return None

nlp = load_nlp()

st.set_page_config(page_title="KAFBOC AI Miner v4", layout="wide")
st.title("📂 KAFBOC Ultra-Smart Data Extractor")

def deep_clean_name(text):
    # 1. Spacing Fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # 2. AI Entity Extraction
    if nlp:
        doc = nlp(text[:1500]) # First 1500 chars focus
        # Sirf wahi entities uthayen jo 'PERSON' hon
        candidate_names = [ent.text.strip() for ent in doc.ents if ent.label_ == "PERSON"]
        
        # Skill/Heading Blocklist (Aapki screenshot ke mutabiq)
        block_list = [
            'communications', 'closing', 'reporting', 'modeling', 'accounting',
            'contact', 'education', 'skills', 'experience', 'summary', 'profile',
            'karachi', 'pakistan', 'lahore', 'certified', 'chartered', 'curriculum',
            'university', 'association', 'about', 'linkedin', 'professional'
        ]

        for name in candidate_names:
            # Clean extra spaces inside name
            clean_n = " ".join(name.split())
            name_low = clean_n.lower()
            
            # Validation Checks
            if 3 < len(clean_n) < 30:
                if not any(word in name_low for word in block_list):
                    if not any(char.isdigit() for char in clean_n):
                        # Name mein kam se kam ek space honi chahiye (First & Last Name)
                        if " " in clean_n:
                            return clean_n.title()
    
    # 3. Fallback (Agar AI fail ho jaye)
    lines = [l.strip() for l in text.split('\n') if l.strip()][:15]
    for line in lines:
        if len(line) < 30 and " " in line and not any(char.isdigit() for char in line):
            if not any(w in line.lower() for w in ['resume', 'cv', 'email', 'phone']):
                return line.title()
                
    return "Check Manual"

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
        
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        return {
            "File Name": f_name,
            "Name": deep_clean_name(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI ---
files = st.file_uploader("Upload Resumes", accept_multiple_files=True)

if files:
    with st.spinner('AI is analyzing documents...'):
        data = [extract_data(f) for f in files]
        df = pd.DataFrame(data)
    
    st.success("Analysis Complete!")
    st.table(df) # Proper alignment ke liye table behtar hai
    
    # Professional Excel Export
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='KAFBOC_Data')
        workbook = writer.book
        worksheet = writer.sheets['KAFBOC_Data']
        
        # Header Styling
        fmt = workbook.add_format({'bold': True, 'bg_color': '#1f4e78', 'font_color': 'white', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Final Excel Report", output.getvalue(), "KAFBOC_Final.xlsx")
