import streamlit as st
import pandas as pd
import random
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION & CUSTOM THEMING
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="TEDDI | Tagged Exam Database",
    page_icon="🧸",
    layout="wide"
)

# Custom CSS for modern styling
st.markdown("""
    <style>
    /* Main Background & Fonts */
    .main {
        background-color: #F8FAFC;
    }
    
    /* Custom Header Banner */
    .teddi-header {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%);
        padding: 1.8rem 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(79, 70, 229, 0.2);
    }
    .teddi-header h1 {
        color: white !important;
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        margin: 0;
        font-size: 2.2rem;
    }
    .teddi-header p {
        color: #E0E7FF !important;
        font-size: 1.05rem;
        margin-top: 0.3rem;
        margin-bottom: 0;
    }
    
    /* Stat / Count Badges */
    .badge-card {
        background-color: white;
        border-radius: 10px;
        padding: 1rem;
        border-left: 5px solid #4F46E5;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        margin-bottom: 1rem;
    }
    
    /* Style Expandable Question Cards */
    .st-emotion-cache-1h993ip, div[data-testid="stExpander"] {
        background-color: white;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        margin-bottom: 0.8rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }
    
    /* Custom Primary Buttons */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s;
    }
    
    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #F1F5F9;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# HEADER BANNER
# -----------------------------------------------------------------------------
st.markdown("""
    <div class="teddi-header">
        <h1>🧸 TEDDI</h1>
        <p><b>Tagged Exam Database for Departmental Instruction</b> — Search, filter, and build custom exams aligned with Bloom's Taxonomy.</p>
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
        
        for idx, row in group.iterrows():
            q_type = row['Question Type']
            b_level = row[BLOOMS_COL]
            
            q_p = doc.add_paragraph()
            q_p.add_run(f"Q{q_num}. [{q_type} | {b_level}]\n").bold = True
            q_p.add_run(f"{row['Question Part Text']}\n")
            
            if row['Photo associated with the exam question if applicable'] != "N/A":
                q_p.add_run(f"[Associated Diagram/Image: {row['Photo associated with the exam question if applicable']}]\n").italic = True
            
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
st.sidebar.title("🛠️ Exam Builder")

basket_count = len(st.session_state.selected_indices)
st.sidebar.metric(label="Questions in Exam Basket", value=basket_count)

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
        
        if row['Photo associated with the exam question if applicable'] != "N/A":
            st.warning(f"🖼️ **Diagram Reference:** {row['Photo associated with the exam question if applicable']}")
        
        st.markdown("---")
        if is_in_basket:
            if st.button("➖ Remove from Exam Basket", key=f"btn_rem_{idx}"):
                st.session_state.selected_indices.remove(idx)
                st.rerun()
        else:
            if st.button("➕ Add to Exam Basket", key=f"btn_add_{idx}"):
                st.session_state.selected_indices.add(idx)
                st.rerun()
