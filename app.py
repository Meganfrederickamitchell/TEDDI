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
    page_icon="\U0001F43B",
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

# Custom CSS: High-Contrast Color Block Palette (Orchid, Turquoise, Orange, Neutral White)
st.markdown("""
    <style>
    /* Neutral White App Canvas */
    .stApp { 
        background-color: #FFFFFF !important; 
    }
    
    /* Header Banner - Solid Orchid Color Block */
    .teddi-header {
        background-color: #9333EA !important;
        padding: 1.8rem 2rem;
        border-radius: 16px;
        color: #FFFFFF !important;
        margin-bottom: 2rem;
        display: flex;
        align-items: center;
        gap: 1.5rem;
        box-shadow: 0 4px 14px rgba(147, 51, 234, 0.25);
    }
    .teddi-header-text h1 {
        color: #FFFFFF !important;
        font-family: 'Inter', system-ui, sans-serif;
        font-weight: 800;
        margin: 0;
        font-size: 2.3rem;
    }
    .teddi-header-text p {
        color: #F3E8FF !important;
        font-size: 1.05rem;
        margin-top: 0.2rem;
        margin-bottom: 0;
    }
    
    /* Sidebar - Turquoise Framed Soft Block */
    section[data-testid="stSidebar"] {
        background-color: #F0F9FF !important;
        border-right: 3px solid #06B6D4 !important;
    }
    
    /* Question Cards - Clean White Card with Left Orchid Accent Border */
    div[data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-left: 6px solid #9333EA !important;
        border-radius: 12px !important;
        margin-bottom: 0.8rem;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-testid="stExpander"]:hover {
        border-color: #06B6D4 !important;
        border-left: 6px solid #06B6D4 !important;
    }
    
    /* Buttons - Solid Dark Orange Block */
    button[kind="primary"], .stButton>button {
        background-color: #FF7F3E !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        transition: background-color 0.2s ease !important;
    }
    button[kind="primary"]:hover, .stButton>button:hover {
        background-color: #E05D1B !important;
    }
    
    /* Headings Accent */
    h3, .stSubheader {
        color: #0F172A !important;
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
    logo_html = '\U0001F43B'

st.markdown(f"""
    <div class="teddi-header">
        <div style="padding: 0;">
            {logo_html}
        </div>
        <div class="teddi-header-text">
            <h1>TEDDIE</h1>
            <p><b>Tagged Exam Database for Departmental Instruction and Evaluation</b> - Fast, intuitive, Bloom's Taxonomy aligned test builder with embedded figure support.</p>
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
        st.error("CSV file not found!")
        return pd.DataFrame()

    try:
        df = pd.read_csv(csv_file, encoding="utf-8-sig")
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(csv_file, encoding="latin1")
        except Exception as e:
            st.error(f"Error loading CSV: {e}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Error loading CSV: {e}")
        return pd.DataFrame()

    for col in df.columns:
        df[col] = df[col].fillna("N/A")
    return df

df = load_data()

if df.empty:
    st.stop()

BLOOMS_COL = "Bloom's Taxonomy Level"
if BLOOMS_COL not in df.columns:
    for c in df.columns:
        if "bloom" in c.lower():
