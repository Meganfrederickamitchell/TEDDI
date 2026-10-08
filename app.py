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

# --- Load Data Safely ---
@st.cache_data
def load_data():
    df = pd.read_csv("exam_questions_extracted.csv")
    
    # Normalize column names
    col_map = {}
    for col in df.columns:
        c_clean = col.strip().lower()
        if 'topic' in c_clean:
            col_map[col] = 'Topic'
        elif 'bloom' in c_clean:
            col_map[col] = 'Blooms Taxonomy Level'
        elif 'question' in c_clean and 'text' in c_clean:
            col_map[col] = 'Question Text'
        elif 'point' in c_clean:
            col_map[col] = 'Points'
        elif 'correct' in c_clean or 'answer' in c_clean:
            col_map[col] = 'Correct Answer'
        elif 'image' in c_clean:
            col_map[col] = 'Image File'
            
    df = df.rename(columns=col_map)
    
    # Ensure standard column names and string types
    if 'Topic' not in df.columns: df['Topic'] = 'General'
    if 'Blooms Taxonomy Level' not in df.columns: df['Blooms Taxonomy Level'] = 'N/A'
    if 'Points' not in df.columns: df['Points'] = '1'
    if 'Question Text' not in df.columns: df['Question Text'] = 'Question text missing'
    
    df['Topic'] = df['Topic'].fillna('General').astype(str)
    df['Blooms Taxonomy Level'] = df['Blooms Taxonomy Level'].fillna('N/A').astype(str)
    df['Question Text'] = df['Question Text'].fillna('').astype(str)
    
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"Error loading CSV file: {e}")
    st.stop()

# --- Sidebar Configuration ---
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

# Button to auto-generate random set based on category counts
if st.sidebar.button("🎲 Generate Random Exam Set"):
    sampled_indices = []
    for b_level, count in blooms_counts.items():
        if count > 0:
            level_df = filtered_df[filtered_df['Blooms Taxonomy Level'] == b_level]
            sampled = level_df.sample(n=count, random_state=random.randint(1, 10000))
            sampled_indices.extend(sampled.index.tolist())
    st.session_state.selected_q_indices = sampled_indices
    st.rerun()

# --- Main Layout ---
col_bank, col_selected = st.columns([1.2, 1])

with col_bank:
    st.subheader(f"📚 Question Bank ({len(filtered_df)})")
    
    for idx, row in filtered_df.iterrows():
        is_selected = idx in st.session_state.selected_q_indices
        q_num = row['Question Number'] if 'Question Number' in row and pd.notnull(row['Question Number']) else idx + 1
        q_text = str(row['Question Text'])
        
        with st.container():
            st.markdown(f"""
            <div>
                <span class="tag tag-points">{row.get('Points', '1')} Points</span>
                <span class="tag tag-topic">Topic: {row.get('Topic', 'General')}</span>
                <span class="tag tag-blooms">Blooms: {row.get('Blooms Taxonomy Level', 'N/A')}</span>
            </div>
            <h4 style="margin-top: 0.6rem; color: #0F172A;">Q{q_num}. {q_text}</h4>
            """, unsafe_allow_html=True)
            
            # Display Image if available
            img_file = row.get('Image File')
            if pd.notnull(img_file) and str(img_file).strip() not in ['', 'nan', 'N/A']:
                img_path = str(img_file).strip()
                if os.path.exists(img_path):
                    st.image(img_path, width=320)
            
            c1, _ = st.columns([1, 4])
            if is_selected:
                if c1.button("Remove", key=f"rem_{idx}"):
                    st.session_state.selected_q_indices.remove(idx)
                    st.rerun()
            else:
                if c1.button("Add +", key=f"add_{idx}"):
                    st.session_state.selected_q_indices.append(idx)
                    st.rerun()
            st.markdown("---")

with col_selected:
    st.subheader(f"📋 Selected Exam Questions ({len(st.session_state.selected_q_indices)})")
    
    if not st.session_state.selected_q_indices:
        st.info("Use the sidebar counts to auto-generate a set, or click **Add +** on questions individually!")
    else:
        selected_df = df.loc[st.session_state.selected_q_indices].copy()
        
        for idx, row in selected_df.iterrows():
            q_num = row['Question Number'] if 'Question Number' in row and pd.notnull(row['Question Number']) else idx + 1
            st.markdown(f"**Q{q_num}:** {row['Question Text']}")
            if st.button("❌ Remove", key=f"sel_rem_{idx}"):
                st.session_state.selected_q_indices.remove(idx)
                st.rerun()
        
        st.markdown("---")
        
        # --- Word Document Generator ---
        def generate_docx(selected_questions, is_answer_key=False):
            doc = docx.Document()
            
            title = doc.add_paragraph()
            heading_text = "BIOLOGY EXAM - ANSWER KEY" if is_answer_key else "BIOLOGY EXAM"
            r = title.add_run(heading_text)
            r.bold = True
            r.font.size = Pt(20)
            r.font.color.rgb = RGBColor(139, 92, 246) # Orchid
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if not is_answer_key:
                p_sub = doc.add_paragraph("Name: ________________________   Date: ______________\n")
                p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for i, (_, row) in enumerate(selected_questions.iterrows(), 1):
                p_q = doc.add_paragraph()
                r_q = p_q.add_run(f"Question {i} ({row.get('Points', '1')} Points):\n")
                r_q.bold = True
                
                # Clean String Extraction (Fixes "Name: 2, dtype: str" cutoff bug)
                q_text_val = str(row['Question Text'])
                p_q.add_run(q_text_val + "\n")
                
                # Choices
                for col in ['Choice A', 'Choice B', 'Choice C', 'Choice D', 'Choice E']:
                    if col in row and pd.notnull(row[col]) and str(row[col]).strip() not in ['', 'nan']:
                        p_q.add_run(f"   [{col[-1]}] {str(row[col])}\n")
                
                # Correct Answer for Answer Key
                if is_answer_key and 'Correct Answer' in row and pd.notnull(row['Correct Answer']):
                    p_ans = doc.add_paragraph()
                    r_ans = p_ans.add_run(f"   --> CORRECT ANSWER: {str(row['Correct Answer'])}")
                    r_ans.bold = True
                    r_ans.font.color.rgb = RGBColor(6, 182, 212) # Turquoise
                
                # Image Embedding
                img_file = row.get('Image File')
                if pd.notnull(img_file) and str(img_file).strip() not in ['', 'nan', 'N/A']:
                    img_path = str(img_file).strip()
                    if os.path.exists(img_path):
                        doc.add_paragraph()
                        doc.add_picture(img_path, width=Inches(3.5))
                        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
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
            file_name="TEDDI_Exam.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        key_docx = generate_docx(selected_df, is_answer_key=True)
        col_ak.download_button(
            label="🔑 Download Answer Key (.docx)",
            data=key_docx,
            file_name="TEDDI_Answer_Key.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
