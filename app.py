import streamlit as st
import fitz  # PyMuPDF
import docx
import re
import pandas as pd

# Page setup
st.set_page_config(page_title="KAFBOC Data Miner", layout="wide")

st.title("📂 KAFBOC Professional Data Extractor")
st.markdown("---")

def clean_extracted_name(text):
    # Headers aur irrelevant words ko filter karne ke liye
    garbage_keywords = ['contact', 'education', 'experience', 'summary', 'profile', 'address', 'phone', 'mobile', 'resume', 'cv']
    
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    for line in lines:
        # Check 1: Line bahut lambi na ho (aksar paragraph hota hai)
        if len(line) > 40: continue
        
        # Check 2: Email ya numbers na hon
        if "@" in line or any(char.isdigit() for char in line[:5]): continue
        
        # Check 3: Garbage keywords na hon
        if any(k in line.lower() for k in garbage_keywords): continue
        
        # Agar saari checks pass ho jayein, to yehi Name hai
        if len(line) > 2:
            return line
            
    return "Not Found"

def extract_info(uploaded_file):
    text = ""
    file_name = uploaded_file.name
    
    try:
        if file_name.endswith('.pdf'):
            bytes_data = uploaded_file.read()
            doc = fitz.open(stream=bytes_data, filetype="pdf")
            for page in doc:
                text += page.get_text()
        elif file_name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
        
        # FIXED Email Regex (Error khatam karne ke liye)
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        emails = re.findall(email_pattern, text)
        
        return {
            "File Name": file_name,
            "Name": clean_extracted_name(text),
            "Email": emails[0] if emails else "Not Found"
        }
    except Exception as e:
        return {"File Name": file_name, "Name": "Error", "Email": "Processing Failed"}

# --- UI Interface ---
uploaded_files = st.file_uploader("Apni Files Upload Karein", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        data = extract_info(f)
        all_data.append(data)
    
    df = pd.DataFrame(all_data)
    
    st.write("### Extraction Results")
    st.dataframe(df, use_container_width=True)
    
    # Excel format (CSV) download
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Download Result", data=csv, file_name="KAFBOC_Data.csv", mime="text/csv")

st.caption("KAFBOC Tech Services - Secure & Private")
