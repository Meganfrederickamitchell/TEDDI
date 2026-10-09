import streamlit as st
import pandas as pd
import random
import os
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="TEDDI | Tagged Exam Database",
    page_icon="🐻",
    layout="wide"
)

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
# HEADER BANNER
# -----------------------------------------------------------------------------
st.markdown("""
    <div class="teddi-header">
        <div style="font-size: 3.5rem;">🐻</div>
        <div class="teddi-header-text">
            <h1>TEDDIE</h1>
            <p><b>Tagged Exam Database for Departmental Instruction and Evaluation</b> — Fast, intuitive, Bloom's Taxonomy aligned test builder with embedded figure support.</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# LOAD DATA (With Safe Path Resolution)
# -----------------------------------------------------------------------------
@st.cache_data
def load_data():
    csv_candidates = [
        "exam_questions_extracted.csv",
        os.path.join(os.getcwd(), "exam_questions_extracted.csv"),
        os.path.join(os.path.dirname(__file__), "exam_questions_extracted.csv") if '__file__' in globals() else ""
    ]
    
    csv_path = None
    for cand in csv_candidates:
        if cand and os.path.exists(cand):
            csv_path = cand
            break
            
    if not csv_path:
        st.error("⚠️ `exam_questions_extracted.csv` was not found in the root directory. Please verify file upload on GitHub.")
        return pd.DataFrame()

    try:
        df = pd.read_csv(csv_path)
        for col in df.columns:
            df[col] = df[col].fillna("N/A")
        return df
    except Exception as e:
        st.error(f"Error reading CSV file: {e}")
        return pd.DataFrame()

df = load_data()

if df.empty:
    st.stop()

# Auto-detect Blooms Taxonomy column name
BLOOMS_COL = "Bloom's Taxonomy Level"
if BLOOMS_COL not in df.columns:
    for c in df.columns:
        if "bloom" in c.lower():
            BLOOMS_COL = c
            break

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# -----------------------------------------------------------------------------
# SESSION STATE SETUP (Exam Basket)
# -----------------------------------------------------------------------------
if "selected_indices" not in st.session_state:
    st.session_state.selected_indices = set()

# Helper function to resolve image path cleanly
def resolve_image_path(raw_img_val):
    if raw_img_val == "N/A" or pd.isnull(raw_img_val):
        return None
    clean_val = str(raw_img_val).strip()
    
    base_filename = os.path.basename(clean_val)
    candidates = [
        clean_val,
        base_filename,
        os.path.join("images", base_filename),
        os.path.join(os.getcwd(), "images", base_filename)
    ]
    
    for candidate in candidates:
        if candidate and os.path.exists(candidate) and os.path.isfile(candidate):
            return candidate
    return None

# -----------------------------------------------------------------------------
# DOCX GENERATION
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
    
    parent_col = "Parent Question Text" if "Parent Question Text" in selected_df.columns else selected_df.columns[0]
    grouped = selected_df.groupby(parent_col, sort=False)

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
            q_type = row.get('Question Type', 'Question')
            b_level = row.get(BLOOMS_COL, 'N/A')
            q_part = row.get('Question Part Text', row.get('Question Text', ''))

            q_p = doc.add_paragraph()
            q_p.add_run(f"Q{q_num}. [{q_type} | {b_level}]\n").bold = True
            q_p.add_run(f"{q_part}\n")

            if include_answers:
                ans_p = doc.add_paragraph()
                ans_p.add_run(f"Correct Answer: {row.get('Answer Details', row.get('Correct Answer', 'N/A'))}").bold = True
            else:
                if q_type in ['Essay', 'Short Answer', 'Draw']:
                    doc.add_paragraph("\n\n" + "_"*81 + "\n" + "_"*81 + "\n")

            doc.add_paragraph()
            q_num += 1

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# SIDEBAR - AUTO-GENERATOR & EXPORT
# -----------------------------------------------------------------------------
st.sidebar.title("🧸 TEDDI Test Builder")

basket_count = len(st.session_state.selected_indices)
st.sidebar.metric(label="Selected Questions in Basket", value=basket_count)

if basket_count > 0:
    if st.sidebar.button("🗑️ Clear Basket", use_container_width=True):
        st.session_state.selected_indices = set()
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🎲 Auto-Generate Exam")

blooms_levels = df[BLOOMS_COL].unique().tolist()
blooms_targets = {}

st.sidebar.caption("Set target question counts per Bloom's level:")
for level in sorted(blooms_levels):
    count_available = len(df[df[BLOOMS_COL] == level])
    blooms_targets[level] = st.sidebar.number_input(
        f"{level} (Max: {count_available})",
        min_value=0,
        max_value=count_available,
        value=0,
        key=f"target_{level}"
    )

if st.sidebar.button("⚡ Auto-Select Questions", type="primary", use_container_width=True):
    new_selection = set()
    for level, target in blooms_targets.items():
        if target > 0:
            level_indices = df[df[BLOOMS_COL] == level].index.tolist()
            sampled = random.sample(level_indices, min(target, len(level_indices)))
            
            for s_idx in sampled:
                parent_val = df.loc[s_idx, "Parent Question Text"] if "Parent Question Text" in df.columns else "N/A"
                if parent_val != "N/A" and len(str(parent_val).strip()) > 0:
                    sibling_indices = df[df["Parent Question Text"] == parent_val].index.tolist()
                    new_selection.update(sibling_indices)
                else:
                    new_selection.add(s_idx)

    st.session_state.selected_indices = new_selection
    st.sidebar.success(f"Selected {len(new_selection)} questions!")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Export Test Documents")

if len(st.session_state.selected_indices) > 0:
    selected_df = df.loc[sorted(list(st.session_state.selected_indices))]

    student_docx = generate_docx(selected_df, include_answers=False)
    st.sidebar.download_button(
        label="📄 Download Student Exam (.docx)",
        data=student_docx,
        file_name="TEDDI_Student_Exam.docx",
        mime=DOCX_MIME,
        use_container_width=True
    )

    key_docx = generate_docx(selected_df, include_answers=True)
    st.sidebar.download_button(
        label="🔑 Download Answer Key (.docx)",
        data=key_docx,
        file_name="TEDDI_Teacher_Answer_Key.docx",
        mime=DOCX_MIME,
        use_container_width=True
    )
else:
    st.sidebar.info("Add questions to the basket to export docx files.")

# -----------------------------------------------------------------------------
# MAIN CONTENT
# -----------------------------------------------------------------------------
st.subheader("🔍 Explore Question Bank")

col1, col2, col3 = st.columns(3)

with col1:
    selected_blooms = st.multiselect("Filter by Bloom's Level", options=sorted(blooms_levels))
with col2:
    q_types = sorted(df["Question Type"].unique().tolist()) if "Question Type" in df.columns else []
    selected_types = st.multiselect("Filter by Question Type", options=q_types)
with col3:
    search_query = st.text_input("Search Text Keywords", value="", placeholder="e.g. Carbon, Heart, Hemoglobin...")

filtered_df = df.copy()

if selected_blooms:
    filtered_df = filtered_df[filtered_df[BLOOMS_COL].isin(selected_blooms)]
if selected_types and "Question Type" in df.columns:
    filtered_df = filtered_df[filtered_df["Question Type"].isin(selected_types)]
if search_query:
    q_col = "Question Part Text" if "Question Part Text" in df.columns else "Question Text"
    p_col = "Parent Question Text" if "Parent Question Text" in df.columns else q_col
    filtered_df = filtered_df[
        filtered_df[q_col].astype(str).str.contains(search_query, case=False) |
        filtered_df[p_col].astype(str).str.contains(search_query, case=False)
    ]

st.markdown(f"**Showing {len(filtered_df)} of {len(df)} total questions**")

for idx, row in filtered_df.iterrows():
    is_in_basket = idx in st.session_state.selected_indices
    b_lvl = row.get(BLOOMS_COL, 'N/A')
    q_typ = row.get('Question Type', 'Question')
    part_nm = row.get('Part Name', f"Q{idx+1}")

    card_title = f"{'✅ IN BASKET | ' if is_in_basket else ''}[{q_typ}] [{b_lvl}] — {part_nm}"

    with st.expander(card
