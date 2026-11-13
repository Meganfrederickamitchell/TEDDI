
import streamlit as st
import pandas as pd
import hashlib
import base64
import random
import math

from pathlib import Path
from io import BytesIO
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


# =========================================================
# 1. CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="TEDDIE | Exam Builder",
    page_icon="🐻",
    layout="wide"
)

ROOT = Path(__file__).resolve().parent
CSV_FILE = ROOT / "exam_questions_extracted.csv"

MAROON = "#881337"

EXAM = "Exam/Course"
TOPIC = "Topic"
BLOOM = "Bloom's Taxonomy Level"
TYPE = "Question Type"
PARENT = "Parent Question Text"
QUESTION = "Question Part Text"
ANSWER = "Answer Details"
NAME = "Part Name"
IMAGE = "Image File"
OPTIONS = "Options"
ID = "Question ID"
GROUP = "Parent ID"

BLOOMS = [
    "Remember", "Understand", "Apply",
    "Analyze", "Evaluate", "Create"
]

REQUIRED = [
    EXAM, TOPIC, BLOOM, TYPE,
    QUESTION, ANSWER, NAME
]

OPTIONAL = [
    PARENT, IMAGE, OPTIONS, ID, GROUP,
    "Instructor", "Points", "Review Status"
]

NON_ASSESSMENT = [
    "Course Feedback",
    "Positive Affirmation"
]

DOCX_MIME = (
    "application/vnd.openxmlformats-"
    "officedocument.wordprocessingml.document"
)


# =========================================================
# 2. STYLING
# =========================================================

st.markdown("""
<style>
.stApp {
    background-color: white;
}

.block-container {
    padding-top: 2.5rem;
    padding-bottom: 2rem;
}

section[data-testid="stSidebar"] {
    background-color: #FFF1F2;
    border-right: 3px solid #881337;
}

.teddie-header {
    background-color: #881337;
    color: white;
    padding: 1.3rem 1.8rem;
    border-radius: 18px;
    display: flex;
    align-items: center;
    gap: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 15px #88133722;
}

.teddie-header h1 {
    color: white !important;
    font-size: 2.5rem;
    margin: 0;
    padding: 0;
}

.teddie-header p {
    color: #FFE4E6;
    margin: 0.4rem 0;
}

.teddie-header img {
    width: 150px;
    max-width: 25%;
    border-radius: 12px;
    object-fit: contain;
}

div[data-testid="stExpander"] {
    border: 1px solid #E9D5DA;
    border-left: 5px solid #881337;
    border-radius: 10px;
}

button[kind="primary"] {
    background-color: #881337;
    color: white;
    border: none;
}

button[kind="primary"]:hover {
    background-color: #701A2E;
    color: white;
}

[data-testid="stMetricValue"] {
    color: #881337;
}

@media (max-width: 600px) {
    .teddie-header {
        padding: 1rem;
        gap: 0.8rem;
    }
    .teddie-header h1 {
        font-size: 1.8rem;
    }
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# 3. HELPER FUNCTIONS
# =========================================================

def resolve_image(filename):
    """Find images within the project image folders."""

    if not filename:
        return None

    # Handle Windows and Unix paths.
    name = str(filename).replace("\\", "/").split("/")[-1]

    for folder in [ROOT / "IMAGES", ROOT / "images", ROOT]:
        path = folder / name

        if path.is_file() and path.suffix.lower() in {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"
        }:
            return path

    return None


def stable_id(*values):
    """Generate a reproducible identifier."""

    text = "\x1f".join(map(str, values))
    return hashlib.sha256(text.encode()).hexdigest()[:24]


def read_csv(source):
    """Read CSV files with common encodings."""

    try:
        return pd.read_csv(
            source, dtype=str, encoding="utf-8-sig"
        )
    except UnicodeDecodeError:
        if hasattr(source, "seek"):
            source.seek(0)

        return pd.read_csv(
            source, dtype=str, encoding="latin1"
        )


def normalize_data(data):
    """Standardize missing values and database columns."""

    data = data.copy()
    data.columns = data.columns.str.strip()

    if "options" in data.columns and OPTIONS not in data:
        data = data.rename(columns={"options": OPTIONS})

    missing = [c for c in REQUIRED if c not in data]

    if missing:
        raise ValueError(
            "Missing columns: " + ", ".join(missing)
        )

    for col in OPTIONAL:
        if col not in data:
            data[col] = ""

    for col in data.columns:
        data[col] = (
            data[col].fillna("").astype(str).str.strip()
            .replace({"N/A": "", "nan": ""})
        )

    data[BLOOM] = data[BLOOM].str.title()
    data[TYPE] = data[TYPE].str.title()

    # Preserve existing IDs; generate IDs for legacy rows.
    for idx, row in data.iterrows():

        if not row[ID]:
            data.at[idx, ID] = stable_id(
                row[EXAM], row[NAME],
                row[PARENT], row[QUESTION]
            )

        if not row[GROUP]:
            data.at[idx, GROUP] = (
                stable_id(row[EXAM], row[PARENT])
                if row[PARENT]
                else data.at[idx, ID]
            )

    return data


@st.cache_data
def load_data(path, modified_time):
    """Reload automatically when the CSV changes."""

    return normalize_data(read_csv(path))


def choices(data, column):
    return sorted(
        value for value in data[column].unique()
        if value
    )


def filter_questions(data, filters, search=""):
    """Shared filtering for Explore and Auto-Generate."""

    for column, selected in filters.items():
        if selected:
            data = data[data[column].isin(selected)]

    if search.strip():
        text = data[
            [NAME, TOPIC, QUESTION, PARENT]
        ].agg(" ".join, axis=1)

        data = data[
            text.str.contains(
                search.strip(),
                case=False,
                regex=False
            )
        ]

    return data


def filter_controls(data, prefix, include_bloom=True):
    """Reusable question filtering widgets."""

    columns = st.columns(4 if include_bloom else 3)

    exams = columns[0].multiselect(
        "Exam / Course",
        choices(data, EXAM),
        key=f"{prefix}_exam"
    )

    topics = columns[1].multiselect(
        "Topic",
        choices(data, TOPIC),
        key=f"{prefix}_topic"
    )

    types = columns[2].multiselect(
        "Question Type",
        choices(data, TYPE),
        key=f"{prefix}_type"
    )

    filters = {
        EXAM: exams,
        TOPIC: topics,
        TYPE: types
    }

    if include_bloom:
        filters[BLOOM] = columns[3].multiselect(
            "Bloom's Level",
            BLOOMS,
            key=f"{prefix}_bloom"
        )

    return filters


def get_options(row):
    """Return separately stored multiple-choice options."""

    return [
        option.strip()
        for option in row[OPTIONS].split(" | ")
        if option.strip()
    ]


def csv_bytes(data):
    return data.to_csv(index=False).encode("utf-8-sig")


# =========================================================
# 4. LOAD DATABASE AND INITIALIZE BASKET
# =========================================================

if not CSV_FILE.exists():
    st.error(f"Database not found: {CSV_FILE.name}")
    st.stop()

try:
    df = load_data(
        str(CSV_FILE),
        CSV_FILE.stat().st_mtime_ns
    )
except Exception as error:
    st.error(f"Unable to load database: {error}")
    st.stop()

if df.empty:
    st.warning("The question bank is empty.")
    st.stop()

if df[ID].duplicated().any():
    st.error(
        "Duplicate Question IDs detected. "
        "Please review the database before continuing."
    )
    st.stop()

if "basket" not in st.session_state:
    st.session_state.basket = set()

# Remove IDs that no longer exist in the database.
st.session_state.basket.intersection_update(df[ID])


def add_questions(ids):
    st.session_state.basket.update(ids)


def remove_questions(ids):
    st.session_state.basket.difference_update(ids)


def selected_questions():
    return df[df[ID].isin(st.session_state.basket)].copy()


# =========================================================
# 5. HEADER AND SIDEBAR
# =========================================================

logo = resolve_image("bear_silhouette.png")

if logo:
    encoded = base64.b64encode(
        logo.read_bytes()
    ).decode()

    logo_html = (
        f'<img src="data:image/png;base64,{encoded}" '
        'alt="TEDDIE mascot">'
    )
else:
    logo_html = '<span style="font-size:4rem">🐻</span>'

st.markdown(f"""
<div class="teddie-header">
    {logo_html}
    <div>
        <p><b>UMass Amherst Biology</b></p>
        <h1>TEDDIE</h1>
        <p>
            Tagged Exam Database for Departmental
            Instruction and Evaluation
        </p>
        <p>
            Explore, build, and evaluate
            biology assessments.
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.title("🐻 TEDDIE")
st.sidebar.caption("Faculty Exam Builder")

page = st.sidebar.radio(
    "Navigation",
    [
        "Explore Question Bank",
        "Auto-Generate Exam",
        "Exam Basket",
        "Manage Database",
        "Help"
    ]
)

st.sidebar.divider()

st.sidebar.metric(
    "Questions in Basket",
    len(st.session_state.basket)
)

st.sidebar.button(
    "Clear Basket",
    on_click=st.session_state.basket.clear,
    disabled=not st.session_state.basket,
    use_container_width=True
)

st.sidebar.divider()

st.sidebar.metric("Database Questions", len(df))
st.sidebar.metric("Source Exams", df[EXAM].nunique())
st.sidebar.metric("Topics", df[TOPIC].nunique())

st.sidebar.caption(
    "Mascot generated via Google Gemini"
)

st.sidebar.caption(
    "Department of Biology · UMass Amherst"
)


# =========================================================
# 6. QUESTION DISPLAY
# =========================================================

def display_question(row, prefix, show_answer=False):
    """Display a question, context, figure and basket controls."""

    qid = row[ID]
    in_basket = qid in st.session_state.basket

    label = (
        f"{'✓ ' if in_basket else ''}"
        f"{row[NAME]} — {row[BLOOM]} | {row[TYPE]}"
    )

    with st.expander(label):

        st.caption(
            f"{row[EXAM]} · {row[TOPIC]}"
        )

        if row[PARENT]:
            st.markdown("**Scenario / Context**")
            st.write(row[PARENT])

        # Include figures shared by the parent group.
        siblings = df[df[GROUP] == row[GROUP]]

        for filename in choices(siblings, IMAGE):
            path = resolve_image(filename)

            if path:
                st.image(str(path), width=450)
            else:
                st.warning(f"Missing figure: {filename}")

        st.markdown("**Question**")
        st.write(row[QUESTION])

        options = get_options(row)

        if options:
            st.markdown("**Options**")

            for option in options:
                st.write(option)

        if show_answer or st.checkbox(
            "Show answer / rubric",
            key=f"{prefix}_answer_{qid}"
        ):
            st.markdown("**Answer / Rubric**")
            st.info(row[ANSWER] or "No answer recorded.")

        c1, c2 = st.columns(2)

        if in_basket:
            c1.button(
                "Remove from Basket",
                key=f"{prefix}_remove_{qid}",
                on_click=remove_questions,
                args=([qid],)
            )
        else:
            c1.button(
                "Add to Basket",
                key=f"{prefix}_add_{qid}",
                on_click=add_questions,
                args=([qid],),
                type="primary"
            )

        if len(siblings) > 1:
            c2.button(
                f"Add Complete Group ({len(siblings)} parts)",
                key=f"{prefix}_group_{qid}",
                on_click=add_questions,
                args=(siblings[ID].tolist(),)
            )


# =========================================================
# 7. EXAM COMPOSITION SUMMARY
# =========================================================

def composition_summary(data):

    if data.empty:
        return

    higher_order = data[BLOOM].isin(
        ["Analyze", "Evaluate", "Create"]
    ).mean() * 100

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Question Parts", len(data))
    c2.metric("Topics", data[TOPIC].nunique())
    c3.metric("Source Exams", data[EXAM].nunique())
    c4.metric("Higher-Order", f"{higher_order:.0f}%")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("**Bloom's Taxonomy Distribution**")

        counts = (
            data[BLOOM].value_counts()
            .reindex(BLOOMS, fill_value=0)
        )

        for level, count in counts.items():
            st.caption(f"{level}: {count}")
            st.progress(float(count / len(data)))

    with c2:
        st.markdown("**Question Types**")

        st.dataframe(
            data[TYPE].value_counts()
            .rename("Count")
            .to_frame(),
            use_container_width=True
        )

        st.markdown("**Topics Covered**")

        st.dataframe(
            data[TOPIC].value_counts()
            .rename("Count")
            .to_frame(),
            use_container_width=True
        )


# =========================================================
# 8. WORD DOCUMENT EXPORT
# =========================================================

@st.cache_data
def generate_docx(
    selected,
    database,
    mode,
    exam_title,
    course,
    instructions,
    image_version
):
    """Create student, answer-key and faculty documents."""

    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)

    title = exam_title

    if mode == "key":
        title += " — Answer Key"
    elif mode == "combined":
        title += " — Faculty Review"

    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if course:
        doc.add_paragraph(course)

    if mode == "student":
        doc.add_paragraph(
            "Name: ______________________________"
            "    SPIRE ID: ______________"
        )
        doc.add_paragraph(instructions)

    elif mode == "combined":
        doc.add_paragraph(
            "Confidential faculty review copy"
        )

    number = 1

    for parent_id, group in selected.groupby(
        GROUP, sort=False
    ):

        context = group.iloc[0][PARENT]

        if mode != "key":

            if context:
                doc.add_paragraph(context)

            # Include all figures associated with the scenario.
            related = database[
                database[GROUP] == parent_id
            ]

            for filename in choices(related, IMAGE):

                path = resolve_image(filename)

                if path:
                    try:
                        doc.add_picture(
                            str(path),
                            width=Inches(4.8)
                        )
                    except Exception:
                        doc.add_paragraph(
                            f"[Unable to insert figure: {filename}]"
                        )
                else:
                    doc.add_paragraph(
                        f"[Missing figure: {filename}]"
                    )

        for _, row in group.iterrows():

            paragraph = doc.add_paragraph()

            paragraph.add_run(
                f"Q{number}. "
            ).bold = True

            if row["Points"]:
                paragraph.add_run(
                    f"({row['Points']} points) "
                )

            if mode != "key":

                paragraph.add_run(row[QUESTION])

                for option in get_options(row):
                    doc.add_paragraph(option)

            if mode in ["key", "combined"]:

                if mode == "key":
                    paragraph.add_run(row[NAME])

                doc.add_paragraph(
                    f"Answer / Rubric: {row[ANSWER]}"
                )

                doc.add_paragraph(
                    f"Bloom's: {row[BLOOM]} | "
                    f"Topic: {row[TOPIC]} | "
                    f"Type: {row[TYPE]}"
                )

            elif row[TYPE] in [
                "Short Answer", "Essay", "Draw",
                "Open Response"
            ]:

                lines = (
                    6 if row[TYPE] in ["Essay", "Draw"]
                    else 3
                )

                for _ in range(lines):
                    doc.add_paragraph("_" * 75)

            doc.add_paragraph()
            number += 1

    buffer = BytesIO()
    doc.save(buffer)

    return buffer.getvalue()


# =========================================================
# 9. EXPLORE QUESTION BANK
# =========================================================

if page == "Explore Question Bank":

    st.header("Explore Question Bank")

    filters = filter_controls(df, "explore")

    search = st.text_input(
        "Search Keywords",
        placeholder="e.g. Hardy-Weinberg, carbon, heart"
    )

    filtered = filter_questions(df, filters, search)

    st.markdown(
        f"**Showing {len(filtered)} "
        f"of {len(df)} question entries**"
    )

    if filtered.empty:
        st.info("No questions match your filters.")

    else:
        c1, c2 = st.columns(2)

        c1.button(
            f"Add All {len(filtered)} Matching Questions",
            on_click=add_questions,
            args=(filtered[ID].tolist(),)
        )

        show_answers = c2.checkbox(
            "Show all answers"
        )

        # Pagination prevents rendering the entire database.
        page_size = st.selectbox(
            "Questions per page",
            [10, 20, 50],
            index=1
        )

        total_pages = math.ceil(
            len(filtered) / page_size
        )

        page_number = st.selectbox(
            "Page",
            range(1, total_pages + 1)
        )

        start = (page_number - 1) * page_size
        visible = filtered.iloc[start:start + page_size]

        for _, row in visible.iterrows():
            display_question(
                row, "explore", show_answers
            )


# =========================================================
# 10. AUTO-GENERATE EXAM
# =========================================================

elif page == "Auto-Generate Exam":

    st.header("Auto-Generate Exam")

    st.caption(
        "Choose topics, source exams and question types. "
        "Then specify the desired number of question "
        "parts at each Bloom's level."
    )

    filters = filter_controls(
        df, "auto", include_bloom=False
    )

    search = st.text_input(
        "Search Keywords (optional)",
        key="auto_search"
    )

    exclude_feedback = st.checkbox(
        "Exclude feedback and affirmation questions",
        value=True
    )

    pool = filter_questions(df, filters, search)

    if exclude_feedback:
        pool = pool[
            ~pool[TOPIC].isin(NON_ASSESSMENT)
        ]

    st.metric("Available Question Parts", len(pool))

    st.subheader("Target Bloom's Distribution")

    targets = {}
    columns = st.columns(3)

    for i, level in enumerate(BLOOMS):

        available = int(
            (pool[BLOOM] == level).sum()
        )

        targets[level] = columns[i % 3].number_input(
            f"{level} (Available: {available})",
            min_value=0,
            max_value=available,
            value=0,
            step=1,
            key=f"target_{level}_{available}"
        )

    total = sum(targets.values())

    st.metric("Requested Questions", total)

    action = st.radio(
        "Basket behavior",
        [
            "Replace existing basket",
            "Add to existing basket"
        ],
        horizontal=True
    )

    seed = st.number_input(
        "Random seed (for reproducible exams)",
        min_value=0,
        max_value=999999,
        value=42,
        step=1
    )

    if st.button(
        "Generate Exam",
        type="primary",
        disabled=total == 0
    ):

        rng = random.Random(seed)
        new_ids = set()

        # Sample each Bloom's level independently.
        for level, target in targets.items():

            candidates = pool.loc[
                pool[BLOOM] == level, ID
            ].tolist()

            new_ids.update(
                rng.sample(candidates, int(target))
            )

        if action == "Replace existing basket":
            st.session_state.basket = new_ids
        else:
            add_questions(new_ids)

        st.session_state.generation_message = (
            f"Generated {len(new_ids)} question parts. "
            "Review your exam in the Exam Basket."
        )

        st.rerun()

    if "generation_message" in st.session_state:
        st.success(
            st.session_state.pop("generation_message")
        )

    st.info(
        "Auto-generation selects individual question "
        "parts to preserve the requested counts. "
        "Shared scenarios and figures are retained. "
        "Review dependencies between question parts "
        "before exporting."
    )

    if st.session_state.basket:
        st.subheader("Current Exam Composition")
        composition_summary(selected_questions())


# =========================================================
# 11. EXAM BASKET AND EXPORTS
# =========================================================

elif page == "Exam Basket":

    st.header("Review Your Exam")

    selected = selected_questions()

    if selected.empty:
        st.info(
            "Your basket is empty. Explore the question "
            "bank or generate an exam to get started."
        )

    else:

        composition_summary(selected)

        st.divider()

        st.subheader("Selected Questions")

        st.dataframe(
            selected[
                [NAME, TOPIC, BLOOM, TYPE, EXAM]
            ],
            hide_index=True,
            use_container_width=True
        )

        # Identify partially selected multipart groups.
        partial_groups = []

        for parent_id, group in selected.groupby(GROUP):

            full_group = df[df[GROUP] == parent_id]

            if len(group) < len(full_group):
                partial_groups.append(
                    f"{group.iloc[0][NAME]}: "
                    f"{len(group)} of "
                    f"{len(full_group)} parts selected"
                )

        if partial_groups:
            with st.expander(
                "Review incomplete multipart questions"
            ):
                st.warning(
                    "Some selected questions belong to "
                    "larger scenarios. Verify that the "
                    "selected parts work independently."
                )

                for message in partial_groups:
                    st.write(message)

        with st.expander("Preview and Edit Questions"):

            for _, row in selected.iterrows():
                display_question(row, "basket")

        st.divider()

        st.subheader("Customize Exam Documents")

        exam_title = st.text_input(
            "Exam Title",
            value="BIOLOGY EXAM"
        )

        course = st.text_input(
            "Course / Semester",
            placeholder="BIOL 151 — Fall 2026"
        )

        instructions = st.text_area(
            "Student Instructions",
            value=(
                "Answer all questions in the "
                "spaces provided."
            )
        )

        # Check figures required by selected groups.
        related = df[
            df[GROUP].isin(selected[GROUP])
        ]

        missing_images = [
            name for name in choices(related, IMAGE)
            if not resolve_image(name)
        ]

        if missing_images:
            st.warning(
                "Missing figures: "
                + ", ".join(missing_images)
            )

        # Invalidate document cache if figures change.
        image_version = tuple(
            (
                name,
                resolve_image(name).stat().st_mtime_ns
            )
            for name in choices(related, IMAGE)
            if resolve_image(name)
        )

        st.subheader("Download Exam Documents")

        exports = [
            (
                "Student Exam",
                "student",
                "TEDDIE_Student_Exam.docx"
            ),
            (
                "Answer Key",
                "key",
                "TEDDIE_Answer_Key.docx"
            ),
            (
                "Faculty Review",
                "combined",
                "TEDDIE_Faculty_Review.docx"
            )
        ]

        columns = st.columns(3)

        for column, (label, mode, filename) in zip(
            columns, exports
        ):

            document = generate_docx(
                selected,
                df,
                mode,
                exam_title,
                course,
                instructions,
                image_version
            )

            column.download_button(
                label=f"Download {label}",
                data=document,
                file_name=filename,
                mime=DOCX_MIME,
                use_container_width=True
            )

        st.download_button(
            "Download Selected Questions as CSV",
            data=csv_bytes(selected),
            file_name="TEDDIE_Selected_Questions.csv",
            mime="text/csv"
        )


# =========================================================
# 12. DATABASE MANAGEMENT
# =========================================================

elif page == "Manage Database":

    st.header("Manage Question Bank")

    st.caption(
        "Import new exams, inspect database quality "
        "and download an updated master CSV."
    )

    st.subheader("Database Overview")

    composition_summary(df)

    st.dataframe(
        df.groupby(EXAM).size()
        .rename("Question Entries")
        .reset_index(),
        hide_index=True,
        use_container_width=True
    )

    st.divider()

    st.subheader("Database Quality Checks")

    missing_images = df[
        df[IMAGE].apply(
            lambda name: bool(name)
            and resolve_image(name) is None
        )
    ]

    missing_answers = df[df[ANSWER] == ""]

    invalid_blooms = df[
        ~df[BLOOM].isin(BLOOMS)
    ]

    duplicates = df[
        df.duplicated(
            subset=[EXAM, PARENT, QUESTION],
            keep=False
        )
    ]

    # These are review flags, not confirmed errors.
    mc_review = df[
        (df[TYPE] == "Multiple Choice")
        & (df[OPTIONS] == "")
    ]

    checks = {
        "Missing Image Files": missing_images,
        "Missing Answers": missing_answers,
        "Invalid Bloom's Levels": invalid_blooms,
        "Potential Duplicate Questions": duplicates,
        "Multiple Choice — Review Option Storage": mc_review
    }

    for label, problems in checks.items():

        with st.expander(
            f"{label}: {len(problems)}"
        ):

            if problems.empty:
                st.success("No issues detected.")

            else:
                st.dataframe(
                    problems[
                        [EXAM, NAME, QUESTION,
                         BLOOM, ANSWER, IMAGE]
                    ],
                    hide_index=True,
                    use_container_width=True
                )

    st.caption(
        "Multiple-choice options may already appear "
        "in the question or parent scenario. "
        "An empty Options field is a review flag, "
        "not necessarily a missing answer choice."
    )

    st.divider()

    st.subheader("Download Database and Template")

    c1, c2 = st.columns(2)

    c1.download_button(
        "Download Master Database",
        data=csv_bytes(df),
        file_name="exam_questions_extracted.csv",
        mime="text/csv"
    )

    template = pd.DataFrame(
        columns=df.columns
    )

    c2.download_button(
        "Download Empty CSV Template",
        data=csv_bytes(template),
        file_name="TEDDIE_Import_Template.csv",
        mime="text/csv"
    )

    st.divider()

    st.subheader("Import Another Exam")

    uploaded = st.file_uploader(
        "Upload a formatted exam CSV",
        type=["csv"]
    )

    if uploaded is not None:

        try:
            incoming = normalize_data(
                read_csv(uploaded)
            )

            errors = []

            if incoming.empty:
                errors.append(
                    "The uploaded CSV is empty."
                )

            for col in REQUIRED:

                count = int(
                    (incoming[col] == "").sum()
                )

                if count:
                    errors.append(
                        f"{col}: {count} missing values"
                    )

            invalid = incoming.loc[
                ~incoming[BLOOM].isin(BLOOMS),
                BLOOM
            ].unique()

            if len(invalid):
                errors.append(
                    "Invalid Bloom's levels: "
                    + ", ".join(invalid)
                )

            if incoming[ID].duplicated().any():
                errors.append(
                    "Duplicate Question IDs "
                    "within uploaded exam."
                )

            existing_ids = incoming[ID].isin(df[ID])

            if existing_ids.any():
                errors.append(
                    f"{existing_ids.sum()} Question IDs "
                    "already exist in the database."
                )

            combined = pd.concat(
                [df, incoming],
                ignore_index=True
            )

            repeated = combined.duplicated(
                subset=[EXAM, PARENT, QUESTION],
                keep=False
            )

            if repeated.iloc[len(df):].any():
                errors.append(
                    "Potential duplicate questions "
                    "detected in uploaded exam."
                )

            st.markdown("**Import Preview**")

            st.dataframe(
                incoming,
                hide_index=True,
                use_container_width=True
            )

            if errors:

                st.error("Import validation failed.")

                for error in errors:
                    st.write(f"• {error}")

            else:

                st.success(
                    f"{len(incoming)} question entries "
                    "passed structural validation."
                )

                composition_summary(incoming)

                new_images = [
                    name
                    for name in choices(incoming, IMAGE)
                    if not resolve_image(name)
                ]

                if new_images:
                    st.warning(
                        "Add these figures to the "
                        "repository IMAGES folder: "
                        + ", ".join(new_images)
                    )

                st.download_button(
                    "Download Updated Master Database",
                    data=csv_bytes(combined),
                    file_name=CSV_FILE.name,
                    mime="text/csv",
                    type="primary"
                )

                st.info(
                    "Replace the master CSV in your "
                    "private GitHub repository with "
                    "this downloaded file. Add any "
                    "new figures to IMAGES. "
                    "Streamlit will then reload the "
                    "updated database."
                )

        except Exception as error:
            st.error(f"Unable to import CSV: {error}")


# =========================================================
# 13. FACULTY HELP PAGE
# =========================================================

elif page == "Help":

    st.header("Using TEDDIE")

    st.markdown("""
    ### Explore Question Bank

    Filter questions by source exam, biological topic,
    Bloom's taxonomy level, question type, or keyword.

    Add individual questions or complete multipart
    groups to the exam basket.

    ### Auto-Generate Exam

    1. Select source exams, topics and question types.
    2. Enter the desired counts for each Bloom's level.
    3. Choose whether to replace or add to your basket.
    4. Generate the exam.
    5. Review the selected questions in the Exam Basket.

    ### Exam Basket

    Review question content and the cognitive distribution
    of the assessment. Remove unwanted questions.

    Customize the title, course and student instructions.

    Download a student exam, answer key or combined
    faculty review document.

    ### Adding a New Exam

    1. Download the CSV template from Manage Database.
    2. Enter one question part per row.
    3. Assign topics, question types and Bloom's levels.
    4. Include answers, rubrics and image filenames.
    5. Upload the CSV and review validation results.
    6. Download the updated master database.
    7. Replace the master CSV in the GitHub repository.
    8. Add new figures to the IMAGES folder.

    ### Multipart Questions

    Give related question parts the same Parent ID.

    For legacy entries without a Parent ID, TEDDIE
    groups questions by source exam and parent text.

    Review partial question groups before exporting.

    ### Faculty Confidentiality

    TEDDIE contains confidential examination materials.
    Keep the GitHub repository and Streamlit application
    private. Share access only with authorized faculty.

    ### About TEDDIE

    Tagged Exam Database for Departmental Instruction
    and Evaluation.

    Department of Biology, University of Massachusetts
    Amherst.
    """)

    st.subheader("Bloom's Taxonomy Reference")

    definitions = {
        "Remember": "Recall facts or definitions.",
        "Understand": "Explain or interpret concepts.",
        "Apply": "Use knowledge or procedures.",
        "Analyze": "Examine relationships and evidence.",
        "Evaluate": "Make and justify judgments.",
        "Create": "Design or produce something new."
    }

    for level, definition in definitions.items():
        st.markdown(f"**{level}:** {definition}")
