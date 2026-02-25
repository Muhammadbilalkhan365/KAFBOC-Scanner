import streamlit as st
import fitz
import docx
import re
import pandas as pd
import spacy
from io import BytesIO

# AI Model Load with caching
@st.cache_resource
def load_nlp():
    try:
        return spacy.load("en_core_web_sm")
    except:
        return None

nlp = load_nlp()

st.set_page_config(page_title="KAFBOC Smart Miner", layout="wide")
st.title("📂 KAFBOC AI Precision Data Extractor")

def get_smart_name(text):
    # 1. Spacing fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # 2. Blocklist (In alfaz wali lines ko reject karein)
    blocklist = [
        'karachi', 'pakistan', 'education', 'skills', 'experience', 'summary', 
        'profile', 'contact', 'address', 'about', 'communications', 'closing', 
        'reporting', 'modeling', 'accounting', 'certified', 'associate', 'manager', 
        'accountant', 'linkedin', 'email', 'phone', 'mobile', 'resume', 'cv', 
        'page', 'objective', 'hobbies', 'projects', 'mehmoodabad', 'expert'
    ]

    lines = [l.strip() for l in text.split('\n') if l.strip()][:25]
    candidates = []

    if nlp:
        doc = nlp("\n".join(lines))
        # AI se dhoondein ke "Person" kahan hai
        for ent in doc.ents:
            if ent.label_ == "PERSON":
                name = " ".join(ent.text.split())
                # Scorer Logic
                score = 0
                if 2 <= len(name.split()) <= 3: score += 50
                if name.isupper(): score += 10
                if not any(b in name.lower() for b in blocklist): score += 40
                if not any(char.isdigit() for char in name): score += 20
                
                if score > 60:
                    candidates.append((name.title(), score))

    # Sort candidates by score
    candidates.sort(key=lambda x: x[1], reverse=True)
    if candidates:
        return candidates[0][0]

    # Fallback: Agar AI fail ho jaye
    for line in lines:
        cleaned = " ".join(line.split())
        if 2 <= len(cleaned.split()) <= 3 and not any(char.isdigit() for char in cleaned):
            if not any(b in cleaned.lower() for b in blocklist) and len(cleaned) < 30:
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
        
        email = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
        return {
            "File Name": f_name,
            "Name": get_smart_name(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI Layout ---
uploaded_files = st.file_uploader("Upload Resumes", accept_multiple_files=True)

if uploaded_files:
    with st.spinner('KAFBOC AI is analyzing documents...'):
        results = [extract_data(f) for f in uploaded_files]
        df = pd.DataFrame(results)
    
    st.subheader("📋 Final Clean Data")
    st.dataframe(df, use_container_width=True) # Dataframe view
    
    # Excel Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        workbook = writer.book
        worksheet = writer.sheets['Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white'})
        for i, col in enumerate(df.columns):
            worksheet.write(0, i, col, header_fmt)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Excel Report", output.getvalue(), "KAFBOC_Final_Report.xlsx")
