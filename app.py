import streamlit as st
import pandas as pd
import random
import os
import base64
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="TEDDIE | Tagged Exam Database",
    page_icon="🐻",
    layout="wide"
)

# Helper function to convert local image to Base64 HTML string
def get_image_base64(img_path):
    if os.path.exists(img_path):
        with open(img_path, "rb") as f:
            data = f.read()
        return f"data:image/png;base64,{base64.b64encode(data).decode()}"
    return None

# Resolve bear silhouette image path
logo_path = None
for candidate in ["IMAGES/bear_silhouette.png", "images/bear_silhouette.png", "bear_silhouette.png"]:
    if os.path.exists(candidate):
        logo_path = candidate
        break

logo_b64 = get_image_base64(logo_path) if logo_path else None

# Custom CSS
st.markdown("""
    <style>
    .stApp { background-color: #F0F9FF; }
    .teddi-header {
        background: linear-gradient(135deg, #BA55D3 0%, #00CED1 100%);
        padding: 1.8rem 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .teddi-header-text h1 {
        color: white !important;
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        margin: 0;
        font-size: 2.3rem;
    }
    .teddi-header-text p {
        color: #F0FDFA !important;
        font-size: 1.05rem;
        margin-top: 0.2rem;
        margin-bottom: 0;
    }
    section[data-testid="stSidebar"] {
        background-color: #E0F2FE;
        border-right: 1px solid #BAE6FD;
    }
    div[data-testid="stExpander"] {
        background-color: white !important;
        border: 1px solid #BAE6FD !important;
        border-radius: 12px !important;
        margin-bottom: 0.8rem;
    }
    button[kind="primary"], .stButton>button {
        background-color: #FF8C00 !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# HEADER BANNER
# -----------------------------------------------------------------------------
if logo_b64:
    logo_html = f'<img src="{logo_b64}" width="65" style="filter: brightness(0) invert(1);">'
else:
    logo_html = '🐻'

st.markdown(f"""
    <div class="teddi-header">
        <div style="padding: 0;">
            {logo_html}
        </div>
        <div class="teddi-header-text">
            <h1>TEDDIE</h1>
            <p><b>Tagged Exam Database for Departmental Instruction and Evaluation</b> — Fast, intuitive, Bloom's Taxonomy aligned test builder with embedded figure support.</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# LOAD DATA
# -----------------------------------------------------------------------------
@st.cache_data
def load_data():
    csv_file = "exam_questions_extracted.csv"
    if not os.path.exists(csv_file):
        st.error("⚠️ CSV file not found!")
        return pd.DataFrame()

    try:
        df = pd.read_csv(csv_file)
        for col in df.columns:
            df[col] = df[col].fillna("N/A")
        return df
    except Exception as e:
        st.error(f"Error loading CSV: {e}")
        return pd.DataFrame()

df = load_data()

if df.empty:
    st.stop()

BLOOMS_COL = "Bloom's Taxonomy Level"
if BLOOMS_COL not in df.columns:
    for c in df.columns:
        if "bloom" in c.lower():
            BLOOMS_COL = c
            break

DOCX_MIME = "application/vnd.openxmlformats-
