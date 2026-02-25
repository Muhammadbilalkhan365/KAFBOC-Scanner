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
    # Be-fizul keywords jo aksar resumes mein pehle aate hain
    garbage_keywords = ['contact', 'education', 'experience', 'summary', 'profile', 'address', 'phone', 'mobile']
    
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    for line in lines:
        # Check karein ke line mein email, digits ya garbage keywords na hon
        if "@" not in line and not any(k in line.lower() for k in garbage_keywords) and len(line) > 2:
            # Agar line mein sirf numbers hain (jaise phone number), to skip karein
            if not re.match(r'^[0-9+-\s]+$', line):
                return line
    return "N/A"

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
        
        # Email Extraction
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        emails = re.findall(email_pattern, text)
        
        return {
            "File Name": file_name,
            "Name": clean_extracted_name(text),
            "Email": emails[0] if emails else "Not Found"
        }
    except Exception as e:
        return {"File Name": file_name, "Name": "Error", "Email": str(e)}

# UI Layout
uploaded_files = st.file_uploader("Apni PDF ya Word files yahan upload karein:", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    progress_bar = st.progress(0)
    
    for i, f in enumerate(uploaded_files):
        data = extract_info(f)
        all_data.append(data)
        progress_bar.progress((i + 1) / len(uploaded_files))
    
    df = pd.DataFrame(all_data)
    
    st.success(f"{len(uploaded_files)} Files process ho gayi hain!")
    st.dataframe(df, use_container_width=True)
    
    # Professional Download Button
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Data as CSV",
        data=csv,
        file_name="KAFBOC_Extracted_Data.csv",
        mime="text/csv",
    )

st.markdown("---")
st.caption("Developed by Muhammad Bilal | KAFBOC Tech Services")
