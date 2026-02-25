import streamlit as st
import fitz  # PyMuPDF
import docx
import re
import pandas as pd

# Page Configuration
st.set_page_config(page_title="KAFBOC Data Miner", layout="centered")

st.title("📂 KAFBOC Document Scanner")
st.subheader("Extract Names & Emails in Seconds")

def extract_info(uploaded_file):
    text = ""
    file_name = uploaded_file.name
    
    if file_name.endswith('.pdf'):
        bytes_data = uploaded_file.read()
        doc = fitz.open(stream=bytes_data, filetype="pdf")
        for page in doc:
            text += page.get_text()
    elif file_name.endswith('.docx'):
        doc = docx.Document(uploaded_file)
        text = "\n".join([para.text for para in doc.paragraphs])
    
    # Logic for Name and Email
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    emails = re.findall(email_pattern, text)
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    name = lines[0] if lines else "N/A"
    
    return {"File Name": file_name, "Name": name, "Email": emails[0] if emails else "Not Found"}

# File Upload Section
uploaded_files = st.file_uploader("Upload PDF or Word Files", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        data = extract_info(f)
        all_data.append(data)
    
    df = pd.DataFrame(all_data)
    
    # Display Table
    st.write("### Extracted Result")
    st.table(df)
    
    # Download Button
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Download Excel (CSV)", data=csv, file_name="Extracted_Data.csv", mime="text/csv")

st.info("Powered by KAFBOC Tech Services")
