import streamlit as st
import pandas as pd

# Load master CSV
@st.cache_data
def load_data():
    df = pd.read_csv("exam_questions_extracted.csv")
    # Ensure default metadata columns exist
    if "topic" not in df.columns:
        df["topic"] = "General Biology"
    if "exam_name" not in df.columns:
        df["exam_name"] = "BIOL 151 Exam 1 (Fall 2026)"
    return df

df = load_data()

# -----------------------------------------------------------------------------
# SIDEBAR: AUTO-GENERATE EXAM WITH TOPIC & EXAM FILTERING
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("🎯 Auto-Generate Exam Filters")

# Exam & Topic Selectors requested by faculty
selected_exams = st.sidebar.multiselect(
    "Filter by Exam / Course",
    options=sorted(df["exam_name"].unique()),
    default=sorted(df["exam_name"].unique())
)

selected_topics = st.sidebar.multiselect(
    "Filter by Topic (e.g., Hardy-Weinberg, Membrane Fluidity)",
    options=sorted(df["topic"].unique()),
    default=[]
)

# Apply filters
filtered_df = df[df["exam_name"].isin(selected_exams)]
if selected_topics:
    filtered_df = filtered_df[filtered_df["topic"].isin(selected_topics)]

# -----------------------------------------------------------------------------
# MAIN EXPLORE TAB: TOPIC SELECTOR
# -----------------------------------------------------------------------------
st.subheader("Explore Question Bank")
col1, col2, col3 = st.columns(3)

with col1:
    topic_filter = st.multiselect("Filter by Topic", options=sorted(df["topic"].unique()))
with col2:
    blooms_filter = st.multiselect("Filter by Bloom's Level", options=sorted(df["blooms_level"].unique()))
with col3:
    type_filter = st.multiselect("Filter by Question Type", options=sorted(df["question_type"].unique()))

# Apply Explore Filters
explore_df = df.copy()
if topic_filter:
    explore_df = explore_df[explore_df["topic"].isin(topic_filter)]
if blooms_filter:
    explore_df = explore_df[explore_df["blooms_level"].isin(blooms_filter)]
if type_filter:
    explore_df = explore_df[explore_df["question_type"].isin(type_filter)]
