# ===========================
# Import Libraries
# ===========================

import streamlit as st
import pickle
import re
import string
import os
import subprocess
import tempfile
import pdfplumber
from docx import Document
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

# ===========================
# Download NLTK Resources
# ===========================

nltk.download('stopwords')
nltk.download('wordnet')

# ===========================
# Load Saved Files
# ===========================

with open("linear_svc_model.pkl", "rb") as file:
    model = pickle.load(file)

with open("tfidf_vectorizer.pkl", "rb") as file:
    tfidf = pickle.load(file)

with open("label_encoder.pkl", "rb") as file:
    label_encoder = pickle.load(file)

# ===========================
# Initialize Lemmatizer
# ===========================

lemmatizer = WordNetLemmatizer()

# ===========================
# Text Preprocessing Function
# ===========================

def preprocess_text(text):

    text = text.lower()

    text = re.sub(r'http\S+|www\S+', '', text)

    text = re.sub(r'\S+@\S+', '', text)

    text = re.sub(r'\d{10}', '', text)

    text = re.sub(r'<.*?>', '', text)

    text = text.translate(str.maketrans('', '', string.punctuation))

    text = re.sub(r'\d+', '', text)

    text = re.sub(r'\s+', ' ', text).strip()

    words = text.split()

    words = [
        lemmatizer.lemmatize(word)
        for word in words
        if word not in stopwords.words("english")
    ]

    return " ".join(words)

# ===========================
# DOC Text Extraction Function
# ===========================

def extract_doc_text(uploaded_file):

    temp_doc_path = ""

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as temp_doc:
            temp_doc.write(uploaded_file.read())
            temp_doc_path = temp_doc.name

        try:
            import win32com.client

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(temp_doc_path)
            text = doc.Content.Text
            doc.Close(False)
            word.Quit()
            return text

        except Exception as e:
            st.error(f"Unable to extract DOC file: {e}")

       
         

    finally:
        if temp_doc_path and os.path.exists(temp_doc_path):
            os.remove(temp_doc_path)

    return ""

# ===========================
# Streamlit UI
# ===========================

st.set_page_config(
    page_title="Resume Classifier",
    page_icon="RC",
    layout="wide",
)

st.markdown(
    """
    <style>
        .stApp {
            background: #f4f7fb;
            color: #111827;
        }

        .block-container {
            max-width: 1120px;
            padding-top: 2rem;
        }

        h1 {
            color: #0f172a;
            font-size: 2.4rem !important;
            font-weight: 800 !important;
            letter-spacing: 0 !important;
        }

        h3 {
            color: #111827;
            font-weight: 800 !important;
        }

        [data-testid="stMetric"] {
            padding: 1rem;
            border: 1px solid rgba(15, 23, 42, 0.08);
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 14px 30px rgba(15, 23, 42, 0.06);
        }

        [data-testid="stMetricLabel"] {
            color: #64748b;
        }

        [data-testid="stMetricValue"] {
            color: #0f172a;
            font-weight: 800;
        }

        [data-testid="stFileUploader"] {
            padding: 1.15rem;
            border: 1px dashed #2563eb;
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 14px 30px rgba(37, 99, 235, 0.08);
        }

        [data-testid="stFileUploader"] section {
            border: 0;
        }

        .dashboard-strip {
            height: 8px;
            margin: 0.5rem 0 1.3rem;
            border-radius: 8px;
            background: linear-gradient(90deg, #2563eb, #0f766e, #f59e0b);
        }

        .upload-card {
            margin-top: 1.2rem;
            padding: 1.15rem 1.2rem;
            border: 1px solid rgba(15, 23, 42, 0.08);
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 14px 30px rgba(15, 23, 42, 0.06);
        }

        .upload-card-title {
            color: #0f172a;
            font-size: 1.1rem;
            font-weight: 800;
        }

        .result-box {
            margin-top: 1.2rem;
            padding: 1.2rem;
            border-left: 6px solid #0f766e;
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 14px 30px rgba(15, 23, 42, 0.06);
        }

        .result-box strong {
            display: block;
            margin-top: 0.25rem;
            font-size: 1.55rem;
            color: #0f766e;
        }

        .result-box span {
            color: #64748b;
            font-size: 0.82rem;
            font-weight: 900;
            text-transform: uppercase;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Resume Classification Dashboard")
st.markdown('<div class="dashboard-strip"></div>', unsafe_allow_html=True)

metric_1, metric_2, metric_3 = st.columns(3)

with metric_1:
    st.metric("Model", "Linear SVC")

with metric_2:
    st.metric("Formats", "PDF / DOC / DOCX / TXT")

with metric_3:
    st.metric("Status", "Ready")

st.markdown(
    '<div class="upload-card"><div class="upload-card-title">Upload Resume</div></div>',
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "Choose a resume file",
    type=["docx", "pdf", "txt", "doc"]
)

# ===========================
# Resume Prediction
# ===========================

if uploaded_file is not None:

    resume_text = ""

    file_type = uploaded_file.name.split(".")[-1].lower()

    # -------------------------
    # DOCX
    # -------------------------
    if file_type == "docx":

        doc = Document(uploaded_file)

        for para in doc.paragraphs:
            resume_text += para.text + "\n"

    # -------------------------
    # PDF
    # -------------------------
    elif file_type == "pdf":

        with pdfplumber.open(uploaded_file) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:
                    resume_text += text + "\n"

    # -------------------------
    # TXT
    # -------------------------
    elif file_type == "txt":

        resume_text = uploaded_file.read().decode("utf-8")

    # -------------------------
    # DOC
    # -------------------------
    elif file_type == "doc":

        resume_text = extract_doc_text(uploaded_file)

        if not resume_text.strip():
            st.error("Unable to read this DOC file. Please try another DOC file or convert it to DOCX.")

            st.stop()

    # -------------------------
    # Prediction
    # -------------------------

    cleaned_resume = preprocess_text(resume_text)

    resume_vector = tfidf.transform([cleaned_resume])

    prediction = model.predict(resume_vector)

    predicted_category = label_encoder.inverse_transform(prediction)

    st.markdown(
        f"""
        <div class="result-box">
            <span>Predicted category</span>
            <strong>{predicted_category[0]}</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )
