import streamlit as st
import pandas as pd
import random
import os
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION & CUSTOM THEMING
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="TEDDI | Tagged Exam Database",
    page_icon="🐻",
    layout="wide"
)

# Base64 string for the bear silhouette logo
BEAR_LOGO_BASE64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAA_SURBVHhe7dBBAYAAAMAgGP3D2oI9XMAGm4YFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAu4sXkAAf0m4Y4wAAAABJRU5ErkJggg=="

# Custom CSS
st.markdown("""
    <style>
    .stApp { background-color: #F0F9FF; }
    .teddi-header {
        background: linear-gradient(135deg, #BA55D3 0%, #00CED1 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 20px -5px rgba(186, 85, 211, 0.3);
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .teddi-header-text h1 {
        color: white !important;
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        margin: 0;
        font-size: 2.5rem;
        text-shadow: 0 2px 4px rgba(0,0,0,0.15);
    }
    .teddi-header-text p {
        color: #F0FDFA !important;
        font-size: 1.1rem;
        margin-top: 0.3rem;
        margin-bottom: 0;
        font-weight: 500;
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
        box-shadow: 0 2px 5px rgba(0,0,0,0.03) !important;
        transition: all 0.2s ease-in-out;
    }
    div[data-testid="stExpander"]:hover {
        border-color: #00CED1 !important;
        box-shadow: 0 4px 12px rgba(0,206,209,0.15) !important;
    }
    button[kind="primary"], .stButton>button {
        background-color: #FF8C00 !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        transition: transform 0.1s ease, background-color 0.2s !important;
    }
    button[kind="primary"]:hover, .stButton>button:hover {
        background-color: #E07B00 !important;
        transform: translateY(-1px);
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# HEADER BANNER (With Embedded Bear Silhouette)
# -----------------------------------------------------------------------------
st.markdown(f"""
    <div class="teddi-header">
        <div style="background: none; padding: 0;">
            <img src="{BEAR_LOGO_BASE64}" width="65" style="filter: brightness(0) invert(1);">
        </div>
        <div class="teddi-header-text">
            <h1>TEDDIE</h1>
            <p><b>Tagged Exam Database for Departmental Instruction and Evaluation </b> — Fast, intuitive, Bloom's Taxonomy aligned test builder with embedded figure support.</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# LOAD DATA
# -----------------------------------------------------------------------------
@st.cache_data
def load_data():
    try:
        df = pd.read_csv("exam_questions_extracted.csv")
        for col in df.columns:
            df[col] = df[col].fillna("N/A")
        return df
    except Exception as e:
        st.error(f"Error loading CSV file: {e}")
        return pd.DataFrame()

df = load_data()

if df.empty:
    st.stop()

BLOOMS_COL = "Bloom's Taxonomy Level"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# -----------------------------------------------------------------------------
# SESSION STATE SETUP (Exam Basket)
# -----------------------------------------------------------------------------
if "selected_indices" not in st.session_state:
    st.session_state.selected_indices = set()

# Helper function to resolve image path correctly without double-folder errors
def resolve_image_path(raw_img_val):
    if raw_img_val == "N/A" or pd.isnull(raw_img_val):
        return None
    clean_val = str(raw_img_val).strip()
    
    candidates = [
        clean_val,
        os.path.basename(clean_val),
        os.path.join("images", os.path.basename(clean_val))
    ]
    
    for candidate in candidates:
        if candidate and os.path.exists(candidate) and os.path.isfile(candidate):
            return candidate
    return None

# -----------------------------------------------------------------------------
# DOCX GENERATION HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def generate_docx(selected_df, include_answers=False):
    doc = Document()

    title_text = "FINAL EXAM - ANSWER KEY" if include_answers else "FINAL EXAM"
    title = doc.add_heading(title_text, level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not include_answers:
        doc.add_paragraph("Name: _______________________\t\tDate: _____________\n")

    doc.add_paragraph("Instructions: Answer all questions cleanly in the spaces provided.\n")

    selected_df = selected_df.sort_index()
    grouped = selected_df.groupby("Parent Question Text", sort=False)

    q_num = 1
    for parent_text, group in grouped:
        if parent_text != "N/A" and len(str(parent_text).strip()) > 0:
            doc.add_heading("Context / Scenario:", level=3)
            p_context = doc.add_paragraph(str(parent_text))
            p_context.runs[0].font.italic = True

        for _, row in group.iterrows():
            img_path = resolve_image_path(row.get("Image File", "N/A"))
            if img_path:
                try:
                    p_img = doc.add_paragraph()
                    p_img.add_run().add_picture(img_path, width=Inches(4.5))
                    p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    break
                except Exception:
                    pass

        for idx, row in group.iterrows():
            q_type = row['Question Type']
            b_level = row[BLOOMS_COL]

            q_p = doc.add_paragraph()
            q_p.add_run(f"Q{q_num}. [{q_type} | {b_level}]\n").bold = True
            q_p.add_run(f"{row['Question Part Text']}\n")

            if include_answers:
                ans_p = doc.add_paragraph()
                ans_p.add_run(f"Correct Answer: {row['Answer Details']}").bold = True
            else:
                if q_type in ['Essay', 'Short Answer', 'Draw']:
                    doc.add_paragraph("\n\n" + "_"*81 + "\n" + "_"*81 + "\n")

            doc
