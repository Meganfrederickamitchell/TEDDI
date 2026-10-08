import streamlit as st
import pandas as pd
import docx
import os
import random
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="TEDDI - Exam Builder",
    page_icon="🧸",
    layout="wide"
)

# Custom Styling: Turquoise, Orange, and Orchid Color Blocks
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #8B5CF6 0%, #06B6D4 100%);
        padding: 2rem;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 2rem;
        color: #FFFFFF;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        color: #FFFFFF !important;
        font-size: 2.8rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    .main-header p {
        color: #F8FAFC !important;
        font-size: 1.1rem;
        margin: 0;
    }
    .tag {
        display: inline-block;
        padding: 0.3rem 0.8rem;
        border-radius: 16px;
        font-size: 0.85rem;
        font-weight: 700;
        margin-right: 0.4rem;
    }
    .tag-points { background-color: #FF7F3E; color: #FFFFFF; }      /* Orange */
    .tag-topic { background-color: #06B6D4; color: #FFFFFF; }       /* Turquoise */
    .tag-blooms { background-color: #8B5CF6; color: #FFFFFF; }      /* Orchid */
</style>
""", unsafe_allow_html=True)

# --- App Header ---
st.markdown("""
<div class="main-header">
    <h1>🧸 TEDDI - Exam Builder</h1>
    <p>Select questions from the bank or build balanced exams by Blooms Taxonomy category.</p>
</div>
""", unsafe_allow_html=True)

# --- Helper Function: Safely extract a single clean string from any cell ---
def get_cell_text(val):
    if isinstance(val, pd.Series):
        valid_vals = [str(x).strip() for x in val.dropna() if str(x).strip() not in ['', 'nan', 'N/A']]
        return valid_vals[0] if valid_vals else ""
    if pd.isnull(val):
        return ""
    str_val = str(val).strip()
    return "" if str_val in ['nan', 'N/A'] else str_val

# --- Helper Function: Locate Image Path safely across Linux & Streamlit Cloud ---
def get_valid_image_path(img_name):
    clean_name = get_cell_text(img_name)
    if not clean_name:
        return None
    
    # Strip leading slashes
    clean_name = clean_name.lstrip('/\\')
    base_filename = os.path.basename(clean_name)
    
    # List candidate paths to check
    candidates = [
        clean_name,
        os.path.join("images", base_filename),
        os.path.join(os.getcwd(), clean_name),
        os.path.join(os.getcwd(), "images", base_filename)
    ]
    
    # Handle jpg vs png extension mismatch
    if base_filename.endswith(".jpg"):
        candidates.append(os.path.join("images", base_filename.replace(".jpg", ".png")))
    elif base_filename.endswith(".png"):
        candidates.append(os.path.join("images", base_filename.replace(".png", ".jpg")))
    
    for path in candidates:
        if path and os.path.exists(path) and os.path.isfile(path):
            return path
    return None

# --- Load Data Safely ---
@st.cache_data
def load_data():
    df = pd.read_csv("exam_questions_extracted.csv")
    
    # Deduplicate column names if CSV has duplicate headers
    cols = []
    counts = {}
    for col in df.columns:
        c_clean = col.strip()
        if c_clean in counts:
            counts[c_clean] += 1
            cols.append(f"{c_clean}_{counts[c_clean]}")
        else:
            counts[c_clean] = 0
            cols.append(c_clean)
    df.columns = cols
    
    # Standardize column mapping
    col_map = {}
    for col in df.columns:
        c_lower = col.lower()
        if 'topic' in c_lower and 'Topic' not in col_map.values():
            col_map[col] = 'Topic'
        elif 'bloom' in c_lower and 'Blooms Taxonomy Level' not in col_map.values():
            col_map[col] = 'Blooms Taxonomy Level'
        elif 'question' in c_lower and 'text' in c_lower and 'Question Text' not in col_map.values():
            col_map[col] = 'Question Text'
        elif 'point' in c_lower and 'Points' not in col_map.values():
            col_map[col] = 'Points'
        elif ('correct' in c_lower or 'answer' in c_lower) and 'Correct Answer' not in col_map.values():
            col_map[col] = 'Correct Answer'
        elif 'image' in c_lower and 'Image File' not in col_map.values():
            col_map[col] = 'Image File'
        elif 'scenario' in c_lower or 'context' in c_lower:
            col_map[col] = 'Context Scenario'
            
    df = df.rename(columns=col_map)
    
    # Ensure fallbacks exist
    if 'Topic' not in df.columns: df['Topic'] = 'General'
    if 'Blooms Taxonomy Level' not in df.columns: df['Blooms Taxonomy Level'] = 'N/A'
    if 'Points' not in df.columns: df['Points'] = '1'
    if 'Question Text' not in df.columns: df['Question Text'] = 'Question text missing'
    if 'Context Scenario' not in df.columns: df['Context Scenario'] = ''
    
    df['Topic'] = df['Topic'].fillna('General').astype(str)
    df['Blooms Taxonomy Level'] = df['Blooms Taxonomy Level'].fillna('N/A').astype(str)
    
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"Error loading CSV file: {e}")
    st.stop()

# --- Sidebar Filters ---
st.sidebar.header("🔍 Exam Criteria & Filters")

# Topic Multiselect
topics = sorted(list(df['Topic'].unique()))
selected_topics = st.sidebar.multiselect("Filter Topics", options=topics, default=topics)

filtered_df = df[df['Topic'].isin(selected_topics)].copy()

st.sidebar.markdown("---")
st.sidebar.subheader("🎯 Questions by Blooms Category")

# Number of questions selector per Blooms Taxonomy Level
blooms_levels = sorted(list(df['Blooms Taxonomy Level'].unique()))
blooms_counts = {}

for b_level in blooms_levels:
    available_in_level = len(filtered_df[filtered_df['Blooms Taxonomy Level'] == b_level])
    blooms_counts[b_level] = st.sidebar.number_input(
        f"{b_level} (Available: {available_in_level})",
        min_value=0,
        max_value=available_in_level,
        value=0,
        step=1
    )

if 'selected_q_indices' not in st.session_state:
    st.session_state.selected_q_indices = []

# Auto-generate random set based on category counts
if st.sidebar.button("🎲 Generate Random Exam Set"):
    sampled_indices = []
    for b_level, count in blooms_counts.items():
        if count > 0:
            level_df = filtered_df[filtered_df['Blooms Taxonomy Level'] == b_level]
            sampled = level_df.sample(n=count, random_state=random.randint(1, 10000))
            sampled_indices.extend(sampled.index.tolist())
    sampled_indices.sort()
    st.session_state.selected_q_indices = sampled_indices
    st.rerun()

# --- Main Layout ---
col_bank, col_selected = st.columns([1.2, 1])

with col_bank:
    st.subheader(f"📚 Question Bank ({len(filtered_df)})")
    
    for idx, row in filtered_df.iterrows():
        is_selected = idx in st.session_state.selected_q_indices
        q_num = get_cell_text(row.get('Question Number')) or str(idx + 1)
        q_text = get_cell_text(row.get('Question Text'))
        scenario_text = get_cell_text(row.get('Context Scenario'))
        pts = get_cell_text(row.get('Points')) or "1"
        top = get_cell_text(row.get('Topic')) or "General"
        blm = get_cell_text(row.get('Blooms Taxonomy Level')) or "N/A"
        
        with st.container():
            st.markdown(f"""
            <div>
                <span class="tag tag-points">{pts} Points</span>
                <span class="tag tag-topic">Topic: {top}</span>
                <span class="tag tag-blooms">Blooms: {blm}</span>
            </div>
            """, unsafe_allow_html=True)
            
            if scenario_text:
                st.markdown(f"**Context / Scenario:** *{scenario_text}*")
            
            # Display Image if available
            img_path = get_valid_image_path(row.get('Image File'))
            if img_path:
                st.image(img_path, width=420)
            
            st.markdown(f"<h4 style='margin-top: 0.4rem; color: #0F172A;'>Q{q_num}. {q_text}</h4>", unsafe_allow_html=True)
            
            c1, _ = st.columns([1, 4])
            if is_selected:
                if c1.button("Remove", key=f"rem_{idx}"):
                    st.session_state.selected_q_indices.remove(idx)
                    st.rerun()
            else:
                if c1.button("Add +", key=f"add_{idx}"):
                    st.session_state.selected_q_indices.append(idx)
                    st.session_state.selected_q_indices.sort()
                    st.rerun()
            st.markdown("---")

with col_selected:
    st.subheader(f"📋 Selected Exam Questions ({len(st.session_state.selected_q_indices)})")
    
    if not st.session_state.selected_q_indices:
        st.info("Use the sidebar counts to auto-generate a set, or click **Add +** on questions individually!")
    else:
        st.session_state.selected_q_indices.sort()
        selected_df = df.loc[st.session_state.selected_q_indices].copy()
        
        for idx, row in selected_df.iterrows():
            q_num = get_cell_text(row.get('Question Number')) or str(idx + 1)
            q_text = get_cell_text(row.get('Question Text'))
            st.markdown(f"**Q{q_num}:** {q_text}")
            if st.button("❌ Remove", key=f"sel_rem_{idx}"):
                st.session_state.selected_q_indices.remove(idx)
                st.rerun()
        
        st.markdown("---")
        
        # --- Word Document Generator ---
        def generate_docx(selected_questions, is_answer_key=False):
            doc = docx.Document()
            
            # Title
            title = doc.add_paragraph()
            heading_text = "FINAL EXAM - ANSWER KEY" if is_answer_key else "FINAL EXAM"
            r = title.add_run(heading_text)
            r.bold = True
            r.font.size = Pt(22)
            r.font.color.rgb = RGBColor(15, 118, 110)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if not is_answer_key:
                p_sub = doc.add_paragraph("Name: ________________________   Date: ______________\n")
                p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                p_inst = doc.add_paragraph("Instructions: Answer all questions cleanly in the spaces provided.\n")
                p_inst.runs[0].font.italic = True
            
            for i, (_, row) in enumerate(selected_questions.iterrows(), 1):
                scenario_text = get_cell_text(row.get('Context Scenario'))
                if scenario_text:
                    p_scen = doc.add_paragraph()
                    r_lbl = p_scen.add_run("Context / Scenario:\n")
                    r_lbl.bold = True
                    r_lbl.font.color.rgb = RGBColor(15, 118, 110)
                    r_txt = p_scen.add_run(scenario_text)
                    r_txt.font.italic = True
                
                # Image Embedding
                img_path = get_valid_image_path(row.get('Image File'))
                if img_path:
                    doc.add_paragraph()
                    doc.add_picture(img_path, width=Inches(4.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                # Question Line
                p_q = doc.add_paragraph()
                pts = get_cell_text(row.get('Points')) or "1"
                blm = get_cell_text(row.get('Blooms Taxonomy Level')) or "Remember"
                
                r_qnum = p_q.add_run(f"Q{i}. [{blm}] ")
                r_qnum.bold = True
                
                q_text_val = get_cell_text(row.get('Question Text'))
                p_q.add_run(q_text_val + "\n")
                
                # Choices
                for choice_col in ['Choice A', 'Choice B', 'Choice C', 'Choice D', 'Choice E']:
                    c_val = get_cell_text(row.get(choice_col))
                    if c_val:
                        p_q.add_run(f"   [{choice_col[-1]}] {c_val}\n")
                
                # Correct Answer for Answer Key
                if is_answer_key:
                    ans_val = get_cell_text(row.get('Correct Answer'))
                    if ans_val:
                        p_ans = doc.add_paragraph()
                        r_ans = p_ans.add_run(f"   --> CORRECT ANSWER: {ans_val}")
                        r_ans.bold = True
                        r_ans.font.color.rgb = RGBColor(6, 182, 212)
                
                doc.add_paragraph("\n")
                
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            return buffer

        # Export Buttons
        col_ex, col_ak = st.columns(2)
        
        exam_docx = generate_docx(selected_df, is_answer_key=False)
        col_ex.download_button(
            label="📄 Download Exam (.docx)",
            data=exam_docx,
            file_name="TEDDI_Final_Exam.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        key_docx = generate_docx(selected_df, is_answer_key=True)
        col_ak.download_button(
            label="🔑 Download Answer Key (.docx)",
            data=key_docx,
            file_name="TEDDI_Answer_Key.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
