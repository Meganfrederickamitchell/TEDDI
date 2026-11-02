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
    if img_path and os.path.exists(img_path):
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

logo_b64 = get_image_base64(logo_path)

# Custom CSS: UMass Amherst Maroon (#881337) & Neutral White Palette
st.markdown("""
    <style>
    .stApp {
        background-color: #FFFFFF !important;
    }
    .teddi-header {
        background-color: #881337 !important;
        padding: 1.8rem 2.2rem;
        border-radius: 18px;
        color: #FFFFFF !important;
        margin-bottom: 2rem;
        display: flex;
        align-items: center;
        gap: 1.5rem;
        box-shadow: 0 4px 16px rgba(136, 19, 55, 0.2);
    }
    .teddi-header-text h1 {
        color: #FFFFFF !important;
        font-family: 'Inter', system-ui, sans-serif;
        font-weight: 800;
        margin: 0;
        font-size: 2.3rem;
    }
    .teddi-header-text p {
        color: #FFE4E6 !important;
        font-size: 1.05rem;
        margin-top: 0.2rem;
        margin-bottom: 0;
    }
    .umass-badge {
        display: inline-block;
        background-color: #FFFFFF;
        color: #881337;
        font-weight: 800;
        font-size: 0.85rem;
        padding: 0.3rem 0.85rem;
        border-radius: 20px;
        margin-bottom: 0.4rem;
        box-shadow: 0 2px 6px rgba(0,0,0,0.12);
        letter-spacing: 0.4px;
    }
    section[data-testid="stSidebar"] {
        background-color: #FFF1F2 !important;
        border-right: 3px solid #881337 !important;
    }
    div[data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border: 1px solid #FFE4E6 !important;
        border-left: 6px solid #881337 !important;
        border-radius: 12px !important;
        margin-bottom: 0.8rem;
        box-shadow: 0 2px 6px rgba(136, 19, 55, 0.05) !important;
    }
    div[data-testid="stExpander"]:hover {
        border-color: #881337 !important;
        border-left: 6px solid #9F1239 !important;
    }
    button[kind="primary"], .stButton>button {
        background-color: #881337 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        transition: background-color 0.2s ease !important;
    }
    button[kind="primary"]:hover, .stButton>button:hover {
        background-color: #701A2E !important;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# HEADER BANNER
# -----------------------------------------------------------------------------
if logo_b64:
    logo_html = f'<img src="{logo_b64}" style="width: auto; height: 210px; max-height: 210px; object-fit: contain; margin: -10px 0; border-radius: 28px;" alt="TEDDIE bear">'
else:
    logo_html = '🐻'

st.markdown(f"""
    <div class="teddi-header">
        <div style="padding: 0;">
            {logo_html}
        </div>
        <div class="teddi-header-text">
            <span class="umass-badge">🧬 UMass Amherst Biology</span>
            <h1>TEDDIE</h1>
            <p><b>Tagged Exam Database for Departmental Instruction and Evaluation</b> - Fast, intuitive, Bloom's Taxonomy & Topic aligned test builder with embedded figure support.</p>
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

# Key Column Definitions
BLOOMS_COL = "Bloom's Taxonomy Level" if "Bloom's Taxonomy Level" in df.columns else "blooms_level"
EXAM_COL = "Exam/Course" if "Exam/Course" in df.columns else "exam_name"
TOPIC_COL = "Topic" if "Topic" in df.columns else "topic"
TYPE_COL = "Question Type" if "Question Type" in df.columns else "question_type"

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# -----------------------------------------------------------------------------
# SESSION STATE SETUP
# -----------------------------------------------------------------------------
if "selected_indices" not in st.session_state:
    st.session_state.selected_indices = set()

# Helper function to resolve image paths across IMAGES folder
def resolve_image_path(raw_img):
    if raw_img == "N/A" or pd.isnull(raw_img) or str(raw_img).strip() == "":
        return None
    clean_val = str(raw_img).strip()
    base_name = os.path.basename(clean_val)

    candidates = [
        clean_val,
        base_name,
        os.path.join("IMAGES", base_name),
        os.path.join("images", base_name),
        os.path.join(os.getcwd(), "IMAGES", base_name)
    ]

    for cand in candidates:
        if cand and os.path.exists(cand) and os.path.isfile(cand):
            return cand
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
        doc.add_paragraph("Name: _______________________\t\tSPIRE ID: _____________\n")

    doc.add_paragraph("Instructions: Answer all questions cleanly in the spaces provided.\n")

    selected_df = selected_df.sort_index()
    parent_col = "Parent Question Text" if "Parent Question Text" in selected_df.columns else selected_df.columns[0]
    grouped = selected_df.groupby(parent_col, sort=False)

    q_num = 1
    for parent_text, group in grouped:
        for _, row in group.iterrows():
            img_path = resolve_image_path(row.get("Image File", row.get("image_path", "N/A")))
            if img_path:
                try:
                    p_img = doc.add_paragraph()
                    p_img.add_run().add_picture(img_path, width=Inches(4.5))
                    p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    break
                except Exception:
                    pass

        if parent_text != "N/A" and len(str(parent_text).strip()) > 0:
            p_context = doc.add_paragraph(str(parent_text))
            p_context.runs[0].font.italic = True

        for idx, row in group.iterrows():
            q_type = row.get(TYPE_COL, 'Question')
            b_level = row.get(BLOOMS_COL, 'N/A')
            q_part = row.get('Question Part Text', row.get('Question Text', row.get('question_text', '')))

            q_p = doc.add_paragraph()
            q_p.add_run(f"Q{q_num}. [{q_type} | {b_level}]\n").bold = True
            q_p.add_run(f"{q_part}\n")

            if include_answers:
                ans_p = doc.add_paragraph()
                ans_val = row.get('Answer Details', row.get('Correct Answer', row.get('correct_answer', 'N/A')))
                ans_p.add_run(f"Correct Answer: {ans_val}").bold = True
            else:
                if q_type in ['Essay', 'Short Answer', 'Draw', 'Open Response']:
                    doc.add_paragraph("\n\n" + "_"*80 + "\n" + "_"*80 + "\n")

            doc.add_paragraph()
            q_num += 1

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# SIDEBAR
# -----------------------------------------------------------------------------
if logo_b64:
    sidebar_img = f'<img src="{logo_b64}" width="48" style="vertical-align: middle; margin-right: 8px;">'
    st.sidebar.markdown(f"### {sidebar_img} TEDDIE Test Builder", unsafe_allow_html=True)
else:
    st.sidebar.markdown("### 🐻 TEDDIE Test Builder")
st.sidebar.caption("Mascot generated via Google Gemini")

basket_count = len(st.session_state.selected_indices)
st.sidebar.metric(label="Questions in Basket", value=basket_count)

if basket_count > 0:
    if st.sidebar.button("Clear Basket", use_container_width=True):
        st.session_state.selected_indices = set()
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("Auto-Generate Exam")

# Multiselect for Exam/Course and Topic requested by faculty
exam_options = sorted([e for e in df[EXAM_COL].unique().tolist() if e != "N/A"]) if EXAM_COL in df.columns else []
selected_exams = st.sidebar.multiselect("Filter by Exam / Course", options=exam_options, default=exam_options)

topic_options = sorted([t for t in df[TOPIC_COL].unique().tolist() if t != "N/A"]) if TOPIC_COL in df.columns else []
selected_topics = st.sidebar.multiselect("Filter by Topic (e.g., Hardy-Weinberg)", options=topic_options)

# Filter pool for auto-generation
auto_pool_df = df.copy()
if selected_exams and EXAM_COL in auto_pool_df.columns:
    auto_pool_df = auto_pool_df[auto_pool_df[EXAM_COL].isin(selected_exams)]
if selected_topics and TOPIC_COL in auto_pool_df.columns:
    auto_pool_df = auto_pool_df[auto_pool_df[TOPIC_COL].isin(selected_topics)]

blooms_levels = sorted([b for b in df[BLOOMS_COL].unique().tolist() if b != "N/A"])
blooms_targets = {}

st.sidebar.caption("Set target question counts per Bloom's level:")
for level in blooms_levels:
    count_available = len(auto_pool_df[auto_pool_df[BLOOMS_COL] == level])
    blooms_targets[level] = st.sidebar.number_input(
        f"{level} (Max: {count_available})",
        min_value=0,
        max_value=count_available,
        value=0,
        key=f"target_{level}"
    )

if st.sidebar.button("Auto-Select Questions", type="primary", use_container_width=True):
    new_selection = set()
    for level, target in blooms_targets.items():
        if target > 0:
            level_indices = auto_pool_df[auto_pool_df[BLOOMS_COL] == level].index.tolist()
            sampled = random.sample(level_indices, min(target, len(level_indices)))

            for s_idx in sampled:
                parent_val = auto_pool_df.loc[s_idx, "Parent Question Text"] if "Parent Question Text" in auto_pool_df.columns else "N/A"
                if parent_val != "N/A" and len(str(parent_val).strip()) > 0:
                    sibling_indices = df[df["Parent Question Text"] == parent_val].index.tolist()
                    new_selection.update(sibling_indices)
                else:
                    new_selection.add(s_idx)

    st.session_state.selected_indices = new_selection
    st.sidebar.success(f"Selected {len(new_selection)} questions!")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("Export Test Documents")

if len(st.session_state.selected_indices) > 0:
    selected_df = df.loc[sorted(list(st.session_state.selected_indices))]

    student_docx = generate_docx(selected_df, include_answers=False)
    st.sidebar.download_button(
        label="Download Student Exam (.docx)",
        data=student_docx,
        file_name="TEDDI_Student_Exam.docx",
        mime=DOCX_MIME,
        use_container_width=True
    )

    key_docx = generate_docx(selected_df, include_answers=True)
    st.sidebar.download_button(
        label="Download Answer Key (.docx)",
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
st.subheader("Explore Question Bank")

col1, col2, col3, col4 = st.columns(4)

with col1:
    selected_topics_exp = st.multiselect("Filter by Topic", options=topic_options)
with col2:
    selected_blooms_exp = st.multiselect("Filter by Bloom's Level", options=blooms_levels)
with col3:
    q_types = sorted([t for t in df[TYPE_COL].unique().tolist() if t != "N/A"]) if TYPE_COL in df.columns else []
    selected_types_exp = st.multiselect("Filter by Question Type", options=q_types)
with col4:
    search_query = st.text_input("Search Keywords", value="", placeholder="e.g. Carbon, Heart...")

filtered_df = df.copy()

if selected_topics_exp and TOPIC_COL in filtered_df.columns:
    filtered_df = filtered_df[filtered_df[TOPIC_COL].isin(selected_topics_exp)]
if selected_blooms_exp and BLOOMS_COL in filtered_df.columns:
    filtered_df = filtered_df[filtered_df[BLOOMS_COL].isin(selected_blooms_exp)]
if selected_types_exp and TYPE_COL in filtered_df.columns:
    filtered_df = filtered_df[filtered_df[TYPE_COL].isin(selected_types_exp)]
if search_query:
    q_col = "Question Part Text" if "Question Part Text" in filtered_df.columns else "question_text"
    p_col = "Parent Question Text" if "Parent Question Text" in filtered_df.columns else q_col
    filtered_df = filtered_df[
        filtered_df[q_col].astype(str).str.contains(search_query, case=False) |
        filtered_df[p_col].astype(str).str.contains(search_query, case=False)
    ]

st.markdown(f"**Showing {len(filtered_df)} of {len(df)} total questions**")

for idx, row in filtered_df.iterrows():
    is_in_basket = idx in st.session_state.selected_indices
    b_lvl = row.get(BLOOMS_COL, 'N/A')
    q_typ = row.get(TYPE_COL, 'Question')
    part_nm = row.get('Part Name', f"Q{idx+1}")
    top_nm = row.get(TOPIC_COL, 'General')

    card_label = f"{'[IN BASKET] ' if is_in_basket else ''}[{q_typ}] [{top_nm}] [{b_lvl}] - {part_nm}"

    with st.expander(card_label):
        parent_txt = row.get('Parent Question Text', 'N/A')
        if parent_txt != "N/A" and str(parent_txt).strip() != "":
            st.markdown(f"**Scenario / Context:**\n> *{parent_txt}*")

        q_txt = row.get('Question Part Text', row.get('Question Text', row.get('question_text', '')))
        ans_txt = row.get('Answer Details', row.get('Correct Answer', row.get('correct_answer', 'N/A')))

        st.markdown(f"**Question:** {q_txt}")

        opts = row.get('options', row.get('Options', 'N/A'))
        if opts != "N/A" and pd.notna(opts) and str(opts).strip() != "":
            st.markdown("**Options:**")
            for o in str(opts).split(" | "):
                st.write(f"- {o}")

        st.markdown(f"**Key / Answer:** `{ans_txt}`")

        img_path = resolve_image_path(row.get("Image File", row.get("image_path", "N/A")))
        if img_path:
            st.image(img_path, width=450)

        st.markdown("---")
        if is_in_basket:
            if st.button("Remove from Exam Basket", key=f"btn_rem_{idx}"):
                st.session_state.selected_indices.remove(idx)
                st.rerun()
        else:
            if st.button("Add to Exam Basket", key=f"btn_add_{idx}"):
                if parent_txt != "N/A" and len(str(parent_txt).strip()) > 0 and "Parent Question Text" in df.columns:
                    sibling_indices = df[df["Parent Question Text"] == parent_txt].index.tolist()
                    st.session_state.selected_indices.update(sibling_indices)
                else:
                    st.session_state.selected_indices.add(idx)
                st.rerun()
