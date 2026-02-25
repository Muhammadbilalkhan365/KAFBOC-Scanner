import streamlit as st
import fitz
import docx
import re
import pandas as pd
import spacy
from io import BytesIO

# NLP Load with caching
@st.cache_resource
def load_nlp():
    try:
        return spacy.load("en_core_web_sm")
    except:
        return None

nlp = load_nlp()

st.set_page_config(page_title="KAFBOC Advanced Miner", layout="wide")
st.title("📂 KAFBOC AI Data Extractor (Professional Build)")

def is_actually_a_name(candidate):
    # Reject agar number ho
    if any(char.isdigit() for char in candidate): return False
    
    # Reject agar common headings hon
    blocklist = [
        'education', 'experience', 'summary', 'profile', 'skills', 'contact',
        'karachi', 'pakistan', 'lahore', 'islamabad', 'address', 'about',
        'communications', 'closing', 'reporting', 'modeling', 'accounting',
        'university', 'college', 'certified', 'associate', 'manager', 'accountant',
        'linkedin', 'email', 'phone', 'mobile', 'curriculum', 'resume', 'page'
    ]
    
    cand_low = candidate.lower()
    if any(word == cand_low or word in cand_low.split() for word in blocklist):
        return False
        
    # Name mein kam se kam 2 alfaz hone chahiye (First & Last Name)
    words = candidate.split()
    if len(words) < 2 or len(words) > 4:
        return False
        
    return True

def deep_clean_logic(text):
    # 1. Fixing weird spacing (A B D U L -> ABDUL)
    text = re.sub(r'(?<=\b[A-Z])\s(?=[A-Z]\b)', '', text)
    
    # 2. Focus on Top 20 lines (Names are usually here)
    lines = [l.strip() for l in text.split('\n') if l.strip()][:20]
    
    # 3. AI NLP Search
    if nlp:
        doc = nlp("\n".join(lines))
        for ent in doc.ents:
            if ent.label_ == "PERSON":
                clean_n = " ".join(ent.text.split())
                if is_actually_a_name(clean_n):
                    return clean_n.title()

    # 4. Pattern Fallback (Agar AI fail ho jaye)
    for line in lines:
        if is_actually_a_name(line):
            return line.title()
            
    return "Not Found"

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
            "Name": deep_clean_logic(text),
            "Email": email[0] if email else "Not Found"
        }
    except:
        return {"File Name": f_name, "Name": "Error", "Email": "Error"}

# --- UI ---
files = st.file_uploader("Upload Resumes (PDF/Word)", accept_multiple_files=True)

if files:
    with st.spinner('KAFBOC AI is scanning documents...'):
        data = [process_file(f) for f in files]
        df = pd.DataFrame(data)
    
    st.subheader("📋 Extraction Results")
    st.dataframe(df, use_container_width=True)
    
    # Excel Logic
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        worksheet.set_column('A:C', 35)

    st.download_button("📥 Download Excel Report", output.getvalue(), "KAFBOC_Final_Report.xlsx")
