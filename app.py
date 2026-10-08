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
    page_icon="🧸",
    layout="wide"
)

# Custom CSS: Orchid, Turquoise, Dark Orange, and Baby Sky Blue
st.markdown("""
    <style>
    /* Main Background - Soft Baby Sky Blue Tint */
    .stApp {
        background-color: #F0F9FF;
    }
    
    /* Header Banner - Orchid & Turquoise Gradient */
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
    .teddi-logo {
        font-size: 3.8rem;
        background: rgba(255, 255, 255, 0.2);
        padding: 0.5rem 1rem;
        border-radius: 20px;
        backdrop-filter: blur(5px);
    }
    
    /* Sidebar Styling - Soft Baby Sky Blue */
    section[data-testid="stSidebar"] {
        background-color: #E0F2FE;
        border-right: 1px solid #BAE6FD;
    }
    
    /* Expandable Cards - Clean White with Turquoise Border on Hover */
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
    
    /* Primary Buttons - Dark Orange */
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
        <div class="teddi-logo">🧸</div>
        <div class="teddi-header-text">
            <h1>TEDDI</h1>
            <p><b>Tagged Exam Database for Departmental Instruction</b> — Fast, intuitive, Bloom's-aligned test builder with embedded figure support.</p>
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

# -----------------------------------------------------------------------------
# SESSION STATE SETUP (Exam Basket)
# -----------------------------------------------------------------------------
if "selected_indices" not in st.session_state:
    st.session_state.selected_indices = set()

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
    
    grouped = selected_df.groupby("Parent Question Text", sort=False)
    
    q_num = 1
    for parent_text, group in grouped:
        if parent_text != "N/A" and len(str(parent_text).strip()) > 0:
            doc.add_heading("Context / Scenario:", level=3)
            p_context = doc.add_paragraph(str(parent_text))
            p_context.runs[0].font.italic = True
            doc.add_paragraph()
        
        # Check if parent group has an associated image
        first_img = group["Image File"].iloc[0] if "Image File" in group.columns else "N/A"
        if first_img != "N/A" and os.path.exists(f"images/{first_img}"):
            try:
                doc.add_paragraph().add_run().add_picture(f"images/{first_img}", width=Inches(4.5))
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
                    doc.add_paragraph("\n\n_________________________________________________________________________________\n" * 2)
            
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
            new_selection.update(sampled)
    st.session_state.selected_indices = new_selection
    st.sidebar.success(f"Selected {len(new_selection)} questions!")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Export Test Documents")

if len(st.session_state.selected_indices) > 0:
    selected_df = df.loc[list(st.session_state.selected_indices)]
    
    student_docx = generate_docx(selected_df, include_answers=False)
    st.sidebar.download_button(
        label="📄 Download Student Exam (.docx)",
        data=student_docx,
        file_name="TEDDI_Student_Exam.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    
    key_docx = generate_docx(selected_df, include_answers=True)
    st.sidebar.download_button(
        label="🔑 Download Answer Key (.docx)",
        data=key_docx,
        file_name="TEDDI_Teacher_Answer_Key.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
else:
    st.sidebar.info("Add questions to the basket to export docx files.")

# -----------------------------------------------------------------------------
# MAIN CONTENT - FILTERS & QUESTION TABLE
# -----------------------------------------------------------------------------
st.subheader("🔍 Explore Question Bank")

col1, col2, col3 = st.columns(3)

with col1:
    selected_blooms = st.multiselect("Filter by Bloom's Level", options=sorted(blooms_levels))
with col2:
    q_types = sorted(df["Question Type"].unique().tolist())
    selected_types = st.multiselect("Filter by Question Type", options=q_types)
with col3:
    search_query = st.text_input("Search Text Keywords", value="", placeholder="e.g. Carbon, Heart, Hemoglobin...")

filtered_df = df.copy()

if selected_blooms:
    filtered_df = filtered_df[filtered_df[BLOOMS_COL].isin(selected_blooms)]
if selected_types:
    filtered_df = filtered_df[filtered_df["Question Type"].isin(selected_types)]
if search_query:
    filtered_df = filtered_df[
        filtered_df["Question Part Text"].str.contains(search_query, case=False) |
        filtered_df["Parent Question Text"].str.contains(search_query, case=False) |
        filtered_df["Part Name"].str.contains(search_query, case=False)
    ]

st.markdown(f"**Showing {len(filtered_df)} of {len(df)} total questions**")

for idx, row in filtered_df.iterrows():
    is_in_basket = idx in st.session_state.selected_indices
    b_lvl = row[BLOOMS_COL]
    q_typ = row['Question Type']
    
    card_title = f"{'✅ IN BASKET | ' if is_in_basket else ''}[{q_typ}] [{b_lvl}] — {row['Part Name']}"
    
    with st.expander(card_title):
        if row['Parent Question Text'] != "N/A":
            st.markdown(f"**📖 Context / Scenario:**\n> *{row['Parent Question Text']}*")
        
        st.markdown(f"**❓ Question:** {row['Question Part Text']}")
        st.markdown(f"**💡 Key / Answer:** `{row['Answer Details']}`")
        
        # Display extracted image if present
        img_file = row.get("Image File", "N/A")
        if img_file != "N/A" and os.path.exists(f"images/{img_file}"):
            st.image(f"images/{img_file}", caption=row["Photo associated with the exam question if applicable"], width=450)
        
        st.markdown("---")
        if is_in_basket:
            if st.button("➖ Remove from Exam Basket", key=f"btn_rem_{idx}"):
                st.session_state.selected_indices.remove(idx)
                st.rerun()
        else:
            if st.button("➕ Add to Exam Basket", key=f"btn_add_{idx}"):
                st.session_state.selected_indices.add(idx)
                st.rerun()
