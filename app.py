import streamlit as st
import pandas as pd
import docx
import os
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- Page Configuration & Styling ---
st.set_page_config(
    page_title="TEDDI - Exam Builder",
    page_icon="🧸",
    layout="wide"
)

# Custom Styling: Warm Mesh Gradient (Orange, Turquoise, Orchid)
st.markdown("""
<style>
    /* Gradient Background mesh: Dark Orange, Turquoise, Orchid, Sky Blue */
    .stApp {
        background: linear-gradient(135deg, #FF7F3E 0%, #22D3EE 40%, #A855F7 75%, #38BDF8 100%) !important;
        background-attachment: fixed !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Main Header Styling */
    .main-header {
        background: rgba(255, 255, 255, 0.92);
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.15);
        text-align: center;
        margin-bottom: 2rem;
        border: 2px solid rgba(255, 255, 255, 0.5);
    }
    
    .main-header h1 {
        color: #C2410C;
        font-size: 2.8rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }
    
    .author-credit {
        color: #0F766E;
        font-size: 1.05rem;
        font-weight: 600;
        margin-top: 0.3rem;
        margin-bottom: 0.8rem;
    }
    
    .main-header p {
        color: #6B21A8;
        font-size: 1.1rem;
        font-weight: 500;
    }
    
    /* Question Cards */
    .question-card {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 1.5rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border-left: 6px solid #FF7F3E;
    }
    
    /* Tags */
    .tag {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-right: 0.5rem;
    }
    .tag-topic { background-color: #E0F2FE; color: #0369A1; }
    .tag-blooms { background-color: #F3E8FF; color: #6B21A8; }
    .tag-points { background-color: #FFEDD5; color: #C2410C; }
    
    /* Buttons */
    .stButton>button {
        background-color: #FF7F3E !important;
        color: white !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        border: none !important;
        box-shadow: 0 4px 10px rgba(255, 127, 62, 0.3) !important;
    }
    .stButton>button:hover {
        background-color: #E05D1B !important;
    }
</style>
""", unsafe_allow_html=True)

# --- App Header ---
st.markdown("""
<div class="main-header">
    <h1>🧸 TEDDI - Exam Builder</h1>
    <div class="author-credit">Created by Megan Mitchell, PhD Candidate in the Organismic and Evolutionary Biology Graduate Program</div>
    <p>Select questions from the question bank to generate custom exams and answer keys instantly.</p>
</div>
""", unsafe_allow_html=True)

# --- Load Data ---
@st.cache_data
def load_data():
    df = pd.read_csv("exam_questions_extracted.csv")
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"Error loading exam_questions_extracted.csv: {e}")
    st.stop()

# --- Sidebar Filters ---
st.sidebar.header("🔍 Filter Question Bank")

# Topic Filter
topics = sorted([t for t in df['Topic'].dropna().unique()])
selected_topics = st.sidebar.multiselect("Select Topics", options=topics, default=topics)

# Blooms Taxonomy Filter
blooms = sorted([b for b in df['Blooms Taxonomy Level'].dropna().unique()])
selected_blooms = st.sidebar.multiselect("Blooms Taxonomy", options=blooms, default=blooms)

# Filter Dataframe
filtered_df = df[
    (df['Topic'].isin(selected_topics)) &
    (df['Blooms Taxonomy Level'].isin(selected_blooms))
]

st.sidebar.markdown("---")
st.sidebar.write(f"**Available Questions:** {len(filtered_df)} / {len(df)}")

# --- Main Layout ---
col_bank, col_selected = st.columns([1.2, 1])

# Initialize Session State for Selected Questions
if 'selected_q_ids' not in st.session_state:
    st.session_state.selected_q_ids = []

with col_bank:
    st.subheader("📚 Question Bank")
    
    for idx, row in filtered_df.iterrows():
        q_id = int(row['Question Number']) if pd.notnull(row['Question Number']) else idx
        is_selected = q_id in st.session_state.selected_q_ids
        
        with st.container():
            st.markdown(f"""
            <div class="question-card">
                <div>
                    <span class="tag tag-points">{row.get('Points', 'N/A')} Points</span>
                    <span class="tag tag-topic">{row.get('Topic', 'General')}</span>
                    <span class="tag tag-blooms">Blooms Taxonomy: {row.get('Blooms Taxonomy Level', 'N/A')}</span>
                </div>
                <h4 style="margin-top: 0.8rem; color: #1E293B;">Q{q_id}. {row['Question Text']}</h4>
            </div>
            """, unsafe_allow_html=True)
            
            # Show Image if available
            img_file = row.get('Image File')
            if pd.notnull(img_file) and str(img_file).strip() != "":
                img_path = str(img_file).strip()
                if os.path.exists(img_path):
                    st.image(img_path, width=350)
            
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
        st.info("Click **Add +** on any question from the bank to start building your exam!")
    else:
        selected_df = df[df['Question Number'].isin(st.session_state.selected_q_ids)].copy()
        
        for idx, row in selected_df.iterrows():
            q_id = int(row['Question Number'])
            st.markdown(f"**Q{q_id}:** {row['Question Text']}")
            if st.button("❌ Remove", key=f"sel_rem_{q_id}"):
                st.session_state.selected_q_ids.remove(q_id)
                st.rerun()
        
        st.markdown("---")
        
        # --- Word Document (.docx) Export ---
        def generate_docx(selected_questions):
            doc = docx.Document()
            
            # Title
            title = doc.add_paragraph()
            r = title.add_run("BIOLOGY EXAM")
            r.bold = True
            r.font.size = Pt(20)
            r.font.color.rgb = RGBColor(194, 65, 12)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            p_sub = doc.add_paragraph("Name: ________________________   Date: ______________\n")
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for i, (_, row) in enumerate(selected_questions.iterrows(), 1):
                p_q = doc.add_paragraph()
                r_q = p_q.add_run(f"Question {i} ({row.get('Points', 'N/A')} Points):\n")
                r_q.bold = True
                p_q.add_run(str(row['Question Text']) + "\n")
                
                # Choices
                for col in ['Choice A', 'Choice B', 'Choice C', 'Choice D', 'Choice E']:
                    if col in row and pd.notnull(row[col]) and str(row[col]).strip() != "":
                        p_q.add_run(f"   [{col[-1]}] {row[col]}\n")
                
                # Image Embedding
                img_file = row.get('Image File')
                if pd.notnull(img_file) and str(img_file).strip() != "":
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

        docx_file = generate_docx(selected_df)
        
        st.download_button(
            label="📥 Download Exam (.docx)",
            data=docx_file,
            file_name="TEDDI_Generated_Exam.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
