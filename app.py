import streamlit as st
import pandas as pd
import docx
import os
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- Page Configuration & High-Contrast Light Theme ---
st.set_page_config(
    page_title="TEDDI - Exam Builder",
    page_icon="🧸",
    layout="wide"
)

# Custom Styling: Light Solid Background with Turquoise, Orange, and Orchid Color Blocks
st.markdown("""
<style>
    /* Light solid overall canvas for maximum contrast */
    .stApp {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
    }
    
    /* Main Header Card (Orchid Block) */
    .main-header {
        background-color: #8B5CF6;
        padding: 2rem;
        border-radius: 16px;
        box-shadow: 0 4px 14px rgba(139, 92, 246, 0.25);
        text-align: center;
        margin-bottom: 2rem;
        color: #FFFFFF;
    }
    
    .main-header h1 {
        color: #FFFFFF !important;
        font-size: 2.6rem;
        font-weight: 800;
        margin-bottom: 0.3rem;
    }
    
    .author-credit {
        color: #F3E8FF !important;
        font-size: 1.05rem;
        font-weight: 600;
        margin-top: 0.2rem;
        margin-bottom: 0.6rem;
    }
    
    .main-header p {
        color: #F8FAFC !important;
        font-size: 1.05rem;
    }
    
    /* Question Card Containers */
    .question-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 1.4rem;
        margin-bottom: 1rem;
        border: 1px solid #E2E8F0;
        border-left: 6px solid #FF7F3E; /* Orange accent border */
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    
    /* Color Block Tags */
    .tag {
        display: inline-block;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-size: 0.82rem;
        font-weight: 700;
        margin-right: 0.4rem;
    }
    .tag-points { background-color: #FF7F3E; color: #FFFFFF; }      /* Orange */
    .tag-topic { background-color: #06B6D4; color: #FFFFFF; }       /* Turquoise */
    .tag-blooms { background-color: #A855F7; color: #FFFFFF; }      /* Orchid */
    
    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0;
    }
    
    /* Buttons */
    .stButton>button {
        background-color: #06B6D4 !important; /* Turquoise */
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: none !important;
    }
    .stButton>button:hover {
        background-color: #0891B2 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- App Header ---
st.markdown("""
<div class="main-header">
    <h1>🧸 TEDDI - Exam Builder</h1>
    <div class="author-credit">Created by Megan Mitchell, PhD Candidate in the Organismic and Evolutionary Biology Graduate Program</div>
    <p>Select questions from the bank to construct custom exams and answer keys instantly.</p>
</div>
""", unsafe_allow_html=True)

# --- Load & Clean Data safely ---
@st.cache_data
def load_data():
    df = pd.read_csv("exam_questions_extracted.csv")
    
    # Normalize column names to avoid KeyErrors
    col_map = {}
    for col in df.columns:
        c_lower = col.strip().lower()
        if 'topic' in c_lower:
            col_map[col] = 'Topic'
        elif 'bloom' in c_lower:
            col_map[col] = 'Blooms Taxonomy Level'
        elif 'question' in c_lower and 'text' in c_lower:
            col_map[col] = 'Question Text'
        elif 'point' in c_lower:
            col_map[col] = 'Points'
        elif 'correct' in c_lower or 'answer' in c_lower:
            col_map[col] = 'Correct Answer'
        elif 'image' in c_lower:
            col_map[col] = 'Image File'
            
    df = df.rename(columns=col_map)
    
    # Ensure mandatory fallback columns exist
    for req_col in ['Topic', 'Blooms Taxonomy Level', 'Question Text', 'Points', 'Correct Answer', 'Image File']:
        if req_col not in df.columns:
            df[req_col] = 'N/A'
            
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"Error loading exam_questions_extracted.csv: {e}")
    st.stop()

# --- Sidebar Filters ---
st.sidebar.header("🔍 Filter Question Bank")

# Topic Filter
topics = sorted([str(t) for t in df['Topic'].dropna().unique() if str(t).strip() != ''])
selected_topics = st.sidebar.multiselect("Select Topics", options=topics, default=topics)

# Blooms Taxonomy Filter
blooms = sorted([str(b) for b in df['Blooms Taxonomy Level'].dropna().unique() if str(b).strip() != ''])
selected_blooms = st.sidebar.multiselect("Blooms Taxonomy Level", options=blooms, default=blooms)

# Filter Dataframe
filtered_df = df[
    (df['Topic'].astype(str).isin(selected_topics)) &
    (df['Blooms Taxonomy Level'].astype(str).isin(selected_blooms))
]

st.sidebar.markdown("---")
st.sidebar.write(f"**Available Questions:** {len(filtered_df)} / {len(df)}")

# --- Main Layout ---
col_bank, col_selected = st.columns([1.2, 1])

if 'selected_q_ids' not in st.session_state:
    st.session_state.selected_q_ids = []

with col_bank:
    st.subheader("📚 Question Bank")
    
    for idx, row in filtered_df.iterrows():
        q_id = int(row['Question Number']) if ('Question Number' in row and pd.notnull(row['Question Number'])) else idx + 1
        is_selected = q_id in st.session_state.selected_q_ids
        
        with st.container():
            st.markdown(f"""
            <div class="question-card">
                <div>
                    <span class="tag tag-points">{row.get('Points', '1')} Points</span>
                    <span class="tag tag-topic">Topic: {row.get('Topic', 'General')}</span>
                    <span class="tag tag-blooms">Blooms: {row.get('Blooms Taxonomy Level', 'N/A')}</span>
                </div>
                <h4 style="margin-top: 0.8rem; color: #0F172A;">Q{q_id}. {row['Question Text']}</h4>
            </div>
            """, unsafe_allow_html=True)
            
            # Display Image if mapping exists
            img_file = row.get('Image File')
            if pd.notnull(img_file) and str(img_file).strip() != "" and str(img_file) != 'N/A':
                img_path = str(img_file).strip()
                if os.path.exists(img_path):
                    st.image(img_path, width=320)
            
            c1, c2 = st.columns([1, 4])
            if is_selected:
                if c1.button("Remove", key=f"rem_{q_id}"):
                    st.session_state.selected_q_ids.remove(q_id)
                    st.rerun()
            else:
                if c1.button("Add +", key=f"add_{q_id}"):
                    st.session_state.selected_q_ids.append(q_id)
                    st.rerun()
            st.markdown("---")

with col_selected:
    st.subheader(f"📋 Selected Exam Questions ({len(st.session_state.selected_q_ids)})")
    
    if not st.session_state.selected_q_ids:
        st.info("Click **Add +** on questions from the bank to build your exam!")
    else:
        # Match selected IDs
        if 'Question Number' in df.columns:
            selected_df = df[df['Question Number'].isin(st.session_state.selected_q_ids)].copy()
        else:
            selected_df = df.iloc[st.session_state.selected_q_ids].copy()
            
        for idx, row in selected_df.iterrows():
            q_id = int(row['Question Number']) if 'Question Number' in row else idx + 1
            st.markdown(f"**Q{q_id}:** {row['Question Text']}")
            if st.button("❌ Remove", key=f"sel_rem_{q_id}"):
                st.session_state.selected_q_ids.remove(q_id)
                st.rerun()
        
        st.markdown("---")
        
        # --- Word Document (.docx) Generator ---
        def generate_docx(selected_questions, is_answer_key=False):
            doc = docx.Document()
            
            title = doc.add_paragraph()
            heading_text = "BIOLOGY EXAM - ANSWER KEY" if is_answer_key else "BIOLOGY EXAM"
            r = title.add_run(heading_text)
            r.bold = True
            r.font.size = Pt(20)
            r.font.color.rgb = RGBColor(139, 92, 246) # Orchid Title
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if not is_answer_key:
                p_sub = doc.add_paragraph("Name: ________________________   Date: ______________\n")
                p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for i, (_, row) in enumerate(selected_questions.iterrows(), 1):
                p_q = doc.add_paragraph()
                r_q = p_q.add_run(f"Question {i} ({row.get('Points', '1')} Points):\n")
                r_q.bold = True
                p_q.add_run(str(row['Question Text']) + "\n")
                
                # Choices
                for col in ['Choice A', 'Choice B', 'Choice C', 'Choice D', 'Choice E']:
                    if col in row and pd.notnull(row[col]) and str(row[col]).strip() != "":
                        p_q.add_run(f"   [{col[-1]}] {row[col]}\n")
                
                # Include Correct Answer for Answer Key
                if is_answer_key and 'Correct Answer' in row and pd.notnull(row['Correct Answer']):
                    p_ans = doc.add_paragraph()
                    r_ans = p_ans.add_run(f"   --> CORRECT ANSWER: {row['Correct Answer']}")
                    r_ans.bold = True
                    r_ans.font.color.rgb = RGBColor(6, 182, 212)
                
                # Image Embedding
                img_file = row.get('Image File')
                if pd.notnull(img_file) and str(img_file).strip() != "" and str(img_file) != 'N/A':
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
