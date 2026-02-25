import streamlit as st
import fitz
import docx
import re
import pandas as pd
import spacy
from io import BytesIO

# AI Model Load with Cache
@st.cache_resource
def load_nlp():
    try:
        return spacy.load("en_core_web_sm")
    except:
        return None

nlp = load_nlp()

st.set_page_config(page_title="KAFBOC Master AI", layout="wide")
st.title("📂 KAFBOC AI Data Extractor (Master Build)")

def score_and_get_name(text):
    # 1. Spacing Fix (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    if not nlp: return "Model Error"
    
    # Pooray text ko scan karein magar shuru ke hisse par focus karein
    doc = nlp(text[:2000])
    candidates = []
    
    # Blocklist for keywords
    blocklist = ['education', 'experience', 'summary', 'skills', 'contact', 'karachi', 'pakistan', 'university', 'college', 'accountant', 'manager', 'communications', 'modeling', 'reporting']

    for ent in doc.ents:
        if ent.label_ == "PERSON":
            name = " ".join(ent.text.split())
            name_low = name.lower()
            
            # Scoring Logic
            score = 0
            if 2 <= len(name.split()) <= 3: score += 50  # Name usually has 2-3 words
            if name.isupper(): score += 10               # Many resumes have names in CAPS
            if not any(word in name_low for word in blocklist): score += 30
            if not any(char.isdigit() for char in name): score += 20
            
            # Penalties
            if len(name) < 3 or len(name) > 30: score -= 100
            if "@" in name or "http" in name: score -= 100
            
            candidates.append((name.title(), score))
    
    # Sort by highest score
    candidates.sort(key=lambda x: x[1], reverse=True)
    
    if candidates and candidates[0][1] > 50:
        return candidates[0][0]
    
    # Fallback: Agar AI fail ho jaye
    lines = [l.strip() for l in text.split('\n') if l.strip()][:10]
    for line in lines:
        if 2 <= len(line.split()) <= 3 and not any(char.isdigit() for char in line):
            if not any(w in line.lower() for w in blocklist):
                return line.title()
                
    return "Not Found"

def extract_info(uploaded_file):
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
            "Name": score_and_get_name(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI ---
files = st.file_uploader("Upload Resumes", accept_multiple_files=True)

if files:
    with st.spinner('AI is calculating scores for names...'):
        results = [extract_info(f) for f in files]
        df = pd.DataFrame(results)
    
    st.table(df)
    
    # Excel Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='KAFBOC_Data')
        workbook = writer.book
        worksheet = writer.sheets['KAFBOC_Data']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
        for i, col in enumerate(df.columns):
            worksheet.write(0, i, col, header_fmt)
        worksheet.set_column('A:C', 40)

    st.download_button("📥 Download AI Scored Excel", output.getvalue(), "KAFBOC_AI_Data.xlsx")
