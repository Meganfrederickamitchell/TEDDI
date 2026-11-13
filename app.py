
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
from figure_review import figure_review


# ============================================================
# 1. CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="TEDDIE | Exam Builder",
    page_icon="🐻",
    layout="wide"
)

ROOT = Path(__file__).resolve().parent
CSV = ROOT / "exam_questions_extracted.csv"

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

LEVELS = [
    "Remember",
    "Understand",
    "Apply",
    "Analyze",
    "Evaluate",
    "Create"
]

REQUIRED = [
    EXAM, TOPIC, BLOOM, TYPE,
    QUESTION, ANSWER, NAME
]

OPTIONAL = [
    PARENT, IMAGE, OPTIONS, ID, GROUP,
    "Instructor", "Points", "Review Status"
]

MIME = (
    "application/vnd.openxmlformats-"
    "officedocument.wordprocessingml.document"
)


# ============================================================
# 2. WEBSITE STYLING
# ============================================================

st.markdown("""
<style>

.block-container {
    padding-top: 2.5rem;
}

section[data-testid="stSidebar"] {
    border-right: 3px solid #881337;
}

.teddie {
    background: #881337;
    color: white;
    padding: 1.3rem 1.8rem;
    border-radius: 18px;
    display: flex;
    align-items: center;
    gap: 1.5rem;
    margin-bottom: 1.5rem;
}

.teddie h1 {
    color: white;
    margin: 0;
    padding: 0;
}

.teddie img {
    width: 150px;
    max-width: 25%;
    border-radius: 12px;
}

[data-testid="stExpander"] {
    border-left: 5px solid #881337;
    border-radius: 10px;
}

button[kind="primary"] {
    background: #881337;
    color: white;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# 3. DATABASE HELPER FUNCTIONS
# ============================================================

def resolve_image(name):

    if not name:
        return None

    name = str(name).replace("\\", "/").split("/")[-1]

    for folder in [
        ROOT / "IMAGES",
        ROOT / "images",
        ROOT
    ]:

        path = folder / name

        if (
            path.is_file()
            and path.suffix.lower() in {
                ".png", ".jpg", ".jpeg",
                ".gif", ".bmp", ".webp"
            }
            and path.resolve().is_relative_to(ROOT)
        ):
            return path

    return None


def stable_id(*values):

    text = "\x1f".join(values)

    return hashlib.sha256(
        text.encode()
    ).hexdigest()[:24]


def read_csv(source):

    try:
        return pd.read_csv(
            source,
            dtype=str,
            encoding="utf-8-sig",
            keep_default_na=False
        )

    except UnicodeDecodeError:

        if hasattr(source, "seek"):
            source.seek(0)

        return pd.read_csv(
            source,
            dtype=str,
            encoding="latin1",
            keep_default_na=False
        )


def normalize(data):

    data = data.copy()

    data.columns = data.columns.str.strip()

    if "options" in data and OPTIONS not in data:
        data = data.rename(
            columns={"options": OPTIONS}
        )

    missing = set(REQUIRED) - set(data.columns)

    if missing:
        raise ValueError(
            "Missing columns: "
            + ", ".join(sorted(missing))
        )

    for col in OPTIONAL:

        if col not in data:
            data[col] = ""

    for col in data:

        data[col] = (
            data[col]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace({"N/A": "", "nan": ""})
        )

    data[BLOOM] = data[BLOOM].str.title()
    data[TYPE] = data[TYPE].str.title()

    for i, row in data.iterrows():

        qid = row[ID] or stable_id(
            row[EXAM],
            row[NAME],
            row[PARENT],
            row[QUESTION]
        )

        data.at[i, ID] = qid

        data.at[i, GROUP] = row[GROUP] or (
            stable_id(row[EXAM], row[PARENT])
            if row[PARENT]
            else qid
        )

    return data


@st.cache_data
def load(path, version):

    return normalize(read_csv(path))


def choices(data, col):

    return sorted(
        v for v in data[col].unique() if v
    )


def csv_bytes(data):

    return data.to_csv(
        index=False
    ).encode("utf-8-sig")


def filter_questions(data, filters, search=""):

    for col, values in filters.items():

        if values:
            data = data[data[col].isin(values)]

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


def controls(data, prefix, bloom=True):

    fields = [
        (EXAM, "Exam / Course"),
        (TOPIC, "Topic"),
        (TYPE, "Question Type")
    ]

    if bloom:
        fields.append((BLOOM, "Bloom's Level"))

    columns = st.columns(len(fields))

    filters = {}

    for column, (field, label) in zip(
        columns, fields
    ):

        filters[field] = column.multiselect(
            label,
            LEVELS if field == BLOOM
            else choices(data, field),
            key=f"{prefix}_{field}"
        )

    return filters


def options(row):

    return [
        s.strip()
        for s in row[OPTIONS].split(" | ")
        if s.strip()
    ]


def missing_report(data):

    missing = data[
        data[IMAGE].map(
            lambda n: bool(n)
            and resolve_image(n) is None
        )
    ]

    report = (
        missing.groupby([EXAM, IMAGE])
        .size()
        .reset_index(name="Affected Questions")
    )

    return missing, report


# ============================================================
# 4. LOAD MASTER DATABASE
# ============================================================

try:

    df = load(
        str(CSV),
        CSV.stat().st_mtime_ns
    )

except Exception as e:

    st.error(f"Unable to load database: {e}")
    st.stop()


if df.empty or df[ID].duplicated().any():

    st.error(
        "Empty database or duplicate Question IDs."
    )

    st.stop()


if "basket" not in st.session_state:

    st.session_state.basket = set()


st.session_state.basket.intersection_update(
    df[ID]
)


# ============================================================
# 5. BASKET FUNCTIONS
# ============================================================

def add(ids):

    st.session_state.basket.update(ids)


def remove(ids):

    st.session_state.basket.difference_update(ids)


def selected():

    return df[
        df[ID].isin(st.session_state.basket)
    ].copy()


def related(row):

    return df[
        (df[EXAM] == row[EXAM])
        & (df[GROUP] == row[GROUP])
    ]


# ============================================================
# 6. HEADER AND SIDEBAR
# ============================================================

logo = resolve_image("bear_silhouette.png")

if logo:

    encoded = base64.b64encode(
        logo.read_bytes()
    ).decode()

    logo_html = (
        '<img alt="TEDDIE mascot" '
        f'src="data:image/png;base64,{encoded}">'
    )

else:

    logo_html = "🐻"


st.markdown(
    f"""
    <div class="teddie">
        {logo_html}
        <div>
            UMass Amherst Biology
            <h1>TEDDIE</h1>
            Tagged Exam Database for Departmental
            Instruction and Evaluation
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


st.sidebar.title("🐻 TEDDIE")

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

st.sidebar.metric(
    "Questions in Basket",
    len(st.session_state.basket)
)

st.sidebar.button(
    "Clear Basket",
    on_click=st.session_state.basket.clear,
    disabled=not st.session_state.basket
)

st.sidebar.divider()

st.sidebar.metric(
    "Database Questions",
    len(df)
)

st.sidebar.metric(
    "Source Exams",
    df[EXAM].nunique()
)

st.sidebar.metric(
    "Topics",
    df[TOPIC].nunique()
)

st.sidebar.caption(
    "Mascot generated via Google Gemini"
)


# ============================================================
# 7. DISPLAY QUESTIONS
# ============================================================

def display(row, prefix, show=False):

    qid = row[ID]

    inside = qid in st.session_state.basket

    siblings = related(row)

    label = (
        f"{'✓ ' if inside else ''}"
        f"{row[NAME]} — "
        f"{row[BLOOM]} | {row[TYPE]}"
    )

    with st.expander(label):

        st.caption(
            f"{row[EXAM]} · {row[TOPIC]}"
        )

        if row[PARENT]:
            st.write(row[PARENT])

        for name in choices(siblings, IMAGE):

            path = resolve_image(name)

            if path:

                st.image(
                    str(path),
                    width=450
                )

            else:

                st.warning(
                    f"Missing figure: {name}"
                )

        st.markdown("**Question**")

        st.write(row[QUESTION])

        for opt in options(row):
            st.write(opt)

        if show or st.checkbox(
            "Show answer / rubric",
            key=f"{prefix}_ans_{qid}"
        ):

            st.info(
                row[ANSWER]
                or "No answer recorded."
            )

        a, b = st.columns(2)

        a.button(
            "Remove from Basket"
            if inside
            else "Add to Basket",
            key=f"{prefix}_pick_{qid}",
            on_click=remove if inside else add,
            args=([qid],)
        )

        if len(siblings) > 1:

            b.button(
                f"Add Complete Group "
                f"({len(siblings)} parts)",
                key=f"{prefix}_group_{qid}",
                on_click=add,
                args=(siblings[ID].tolist(),)
            )


# ============================================================
# 8. DATABASE AND EXAM SUMMARY
# ============================================================

def summary(data):

    if data.empty:
        return

    a, b, c, d = st.columns(4)

    a.metric(
        "Question Parts",
        len(data)
    )

    b.metric(
        "Topics",
        data[TOPIC].nunique()
    )

    c.metric(
        "Source Exams",
        data[EXAM].nunique()
    )

    higher = (
        100 * data[BLOOM]
        .isin(LEVELS[3:])
        .mean()
    )

    d.metric(
        "Higher-order",
        f"{higher:.0f}%"
    )

    a, b = st.columns(2)

    with a:

        st.markdown("**Bloom's distribution**")

        counts = (
            data[BLOOM]
            .value_counts()
            .reindex(LEVELS, fill_value=0)
        )

        for level, count in counts.items():

            st.caption(f"{level}: {count}")

            st.progress(
                float(count / len(data))
            )

    with b:

        for col in [TYPE, TOPIC]:

            st.markdown(f"**{col}**")

            st.dataframe(
                data[col]
                .value_counts()
                .rename("Count")
                .to_frame()
            )


# ============================================================
# 9. WORD DOCUMENT EXPORT
# ============================================================

@st.cache_data
def export(
    data,
    database,
    mode,
    title,
    course,
    instructions,
    image_version
):

    doc = Document()

    suffix = {
        "student": "",
        "key": " — Answer Key",
        "combined": " — Faculty Review"
    }

    heading = doc.add_heading(
        title + suffix[mode],
        0
    )

    heading.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    if course:
        doc.add_paragraph(course)

    if mode == "student":

        doc.add_paragraph(
            "Name: _____________________ "
            "   SPIRE ID: _____________"
        )

        doc.add_paragraph(instructions)

    else:

        doc.add_paragraph(
            "Confidential faculty copy"
        )

    n = 1

    for (exam, group_id), group in data.groupby(
        [EXAM, GROUP],
        sort=False
    ):

        if mode != "key":

            if group.iloc[0][PARENT]:

                doc.add_paragraph(
                    group.iloc[0][PARENT]
                )

            siblings = database[
                (database[EXAM] == exam)
                & (database[GROUP] == group_id)
            ]

            for name in choices(siblings, IMAGE):

                path = resolve_image(name)

                if not path:

                    raise ValueError(
                        f"Missing figure: {name}"
                    )

                doc.add_picture(
                    str(path),
                    width=Inches(4.8)
                )

        for _, row in group.iterrows():

            p = doc.add_paragraph()

            p.add_run(f"Q{n}. ").bold = True

            if row["Points"]:

                p.add_run(
                    f"({row['Points']} points) "
                )

            p.add_run(
                row[NAME]
                if mode == "key"
                else row[QUESTION]
            )

            if mode != "key":

                for opt in options(row):

                    doc.add_paragraph(opt)

            if mode != "student":

                doc.add_paragraph(
                    "Answer / Rubric: "
                    + row[ANSWER]
                )

                doc.add_paragraph(
                    f"Bloom's: {row[BLOOM]} | "
                    f"Topic: {row[TOPIC]}"
                )

            elif row[TYPE] in [
                "Short Answer",
                "Essay",
                "Draw",
                "Open Response"
            ]:

                lines = (
                    6 if row[TYPE] in [
                        "Essay", "Draw"
                    ]
                    else 3
                )

                for _ in range(lines):

                    doc.add_paragraph("_" * 75)

            n += 1

    buf = BytesIO()

    doc.save(buf)

    return buf.getvalue()


# ============================================================
# 10. EXPLORE QUESTION BANK
# ============================================================

if page == "Explore Question Bank":

    st.header(page)

    filters = controls(df, "explore")

    search = st.text_input(
        "Search Keywords"
    )

    pool = filter_questions(
        df, filters, search
    )

    st.write(
        f"Showing {len(pool)} "
        f"of {len(df)} question entries"
    )

    if pool.empty:

        st.info("No matching questions.")

    else:

        st.button(
            "Add All Matching Questions",
            on_click=add,
            args=(pool[ID].tolist(),)
        )

        show = st.checkbox(
            "Show all answers"
        )

        size = st.selectbox(
            "Questions per page",
            [10, 20, 50],
            index=1
        )

        pages = math.ceil(
            len(pool) / size
        )

        number = st.selectbox(
            "Page",
            range(1, pages + 1),
            key=f"page_{pages}_{size}"
        )

        for _, row in pool.iloc[
            (number - 1) * size:number * size
        ].iterrows():

            display(
                row,
                "explore",
                show
            )


# ============================================================
# 11. AUTO-GENERATE EXAM
# ============================================================

elif page == "Auto-Generate Exam":

    st.header(page)

    filters = controls(
        df,
        "auto",
        bloom=False
    )

    search = st.text_input(
        "Search Keywords (optional)"
    )

    pool = filter_questions(
        df, filters, search
    )

    if st.checkbox(
        "Exclude feedback and affirmation questions",
        value=True
    ):

        pool = pool[
            ~pool[TOPIC].isin([
                "Course Feedback",
                "Positive Affirmation"
            ])
        ]

    append = st.radio(
        "Basket behavior",
        [
            "Replace existing basket",
            "Add to existing basket"
        ]
    ) == "Add to existing basket"

    if append:

        pool = pool[
            ~pool[ID].isin(
                st.session_state.basket
            )
        ]

    st.metric(
        "Available Question Parts",
        len(pool)
    )

    targets = {}

    cols = st.columns(3)

    for i, level in enumerate(LEVELS):

        n = int(
            (pool[BLOOM] == level).sum()
        )

        targets[level] = cols[
            i % 3
        ].number_input(
            f"{level} (Available: {n})",
            min_value=0,
            max_value=n,
            value=0,
            key=f"target_{level}_{n}"
        )

    seed = st.number_input(
        "Random seed",
        0,
        999999,
        42
    )

    total = sum(targets.values())

    st.metric(
        "Requested Questions",
        total
    )

    if st.button(
        "Generate Exam",
        type="primary",
        disabled=not total
    ):

        rng = random.Random(seed)

        ids = set()

        for level, n in targets.items():

            available = sorted(
                pool.loc[
                    pool[BLOOM] == level,
                    ID
                ]
            )

            ids.update(
                rng.sample(available, n)
            )

        if not append:

            st.session_state.basket.clear()

        add(ids)

        st.session_state.notice = (
            f"Added {len(ids)} question parts."
        )

        st.rerun()

    if "notice" in st.session_state:

        st.success(
            st.session_state.pop("notice")
        )

    st.info(
        "Individual parts are sampled. "
        "Review multipart dependencies before export."
    )

    summary(selected())


# ============================================================
# 12. EXAM BASKET
# ============================================================

elif page == "Exam Basket":

    st.header(page)

    data = selected()

    if data.empty:

        st.info("Your basket is empty.")

    else:

        summary(data)

        st.dataframe(
            data[
                [NAME, TOPIC, BLOOM, TYPE, EXAM]
            ],
            hide_index=True
        )

        partial = []

        for _, group in data.groupby(
            [EXAM, GROUP]
        ):

            if len(group) < len(
                related(group.iloc[0])
            ):

                partial.append(
                    group.iloc[0][NAME]
                )

        if partial:

            st.warning(
                "Review incomplete multipart groups: "
                + ", ".join(partial)
            )

        for _, row in data.iterrows():

            display(row, "basket")

        title = st.text_input(
            "Exam Title",
            "BIOLOGY EXAM"
        )

        course = st.text_input(
            "Course / Semester"
        )

        instructions = st.text_area(
            "Student Instructions",
            "Answer all questions in the spaces provided."
        )

        groups = set(
            zip(data[EXAM], data[GROUP])
        )

        siblings = df[
            [
                pair in groups
                for pair in zip(
                    df[EXAM], df[GROUP]
                )
            ]
        ]

        missing, report = missing_report(
            siblings
        )

        allow_draft = False

        if not missing.empty:

            st.warning(
                "Some selected questions require missing figures. "
                "Resolve these figures before distributing the exam."
            )

            st.dataframe(
                report,
                hide_index=True
            )

            allow_draft = st.checkbox(
                "Allow draft export with missing figures",
                value=False,
                key="allow_draft_export",
                help="Download an incomplete exam for editing and review."
            )

            if allow_draft:
                st.info(
                    "Draft export enabled. Insert missing figures "
                    "before distributing this exam to students."
                )

        version = tuple(
            (
                name,
                resolve_image(name)
                .stat()
                .st_mtime_ns
            )
            for name in choices(siblings, IMAGE)
            if resolve_image(name)
        )

        exports = [
            ("Student Exam", "student"),
            ("Answer Key", "key"),
            ("Faculty Review", "combined")
        ]

        for col, (label, mode) in zip(
            st.columns(3),
            exports
        ):

            if (
                mode != "key"
                and not missing.empty
                and not allow_draft
            ):
                col.button(
                    f"Download {label}",
                    disabled=True,
                    key=f"blocked_export_{mode}",
                    help="Resolve missing figures or enable draft export."
                )
                continue

            try:

                content = export(
                    data,
                    df,
                    mode,
                    title,
                    course,
                    instructions,
                    version
                )

                col.download_button(
                    "Download " + label,
                    content,
                    f"TEDDIE_{mode}.docx",
                    MIME
                )

            except Exception as e:

                col.error(
                    f"Export failed: {e}"
                )

        st.download_button(
            "Download Selected Questions CSV",
            csv_bytes(data),
            "TEDDIE_Selected_Questions.csv",
            "text/csv"
        )


# ============================================================
# 13. MANAGE DATABASE
# ============================================================

elif page == "Manage Database":

    st.header("Manage Question Bank")

    summary(df)

    st.subheader("Database Quality Checks")

    missing, report = missing_report(df)

    with st.expander(
        f"Missing Image Files: "
        f"{missing[IMAGE].nunique()} unique files "
        f"({len(missing)} affected questions)"
    ):

        if missing.empty:

            st.success(
                "All referenced images are available."
            )

        else:

            st.dataframe(
                report,
                hide_index=True
            )

            st.download_button(
                "Download Missing Image Report",
                csv_bytes(report),
                "missing_images_report.csv",
                "text/csv"
            )

    checks = {

        "Missing Answers":
            df[df[ANSWER] == ""],

        "Invalid Bloom's Levels":
            df[~df[BLOOM].isin(LEVELS)],

        "Potential Duplicate Questions":
            df[
                df.duplicated(
                    [EXAM, PARENT, QUESTION],
                    keep=False
                )
            ],

        "Multiple Choice — Review Option Storage":
            df[
                (df[TYPE] == "Multiple Choice")
                & (df[OPTIONS] == "")
            ]
    }

    for label, rows in checks.items():

        with st.expander(
            f"{label}: {len(rows)}"
        ):

            if rows.empty:

                st.success(
                    "No issues detected."
                )

            else:

                st.dataframe(
                    rows[
                        [
                            EXAM,
                            NAME,
                            QUESTION,
                            BLOOM,
                            ANSWER,
                            IMAGE
                        ]
                    ],
                    hide_index=True
                )

    st.caption(
        "Empty Options may mean choices are embedded "
        "in the question or scenario; review against "
        "the original."
    )

    # NEW: FIGURE EXTRACTION AND REVIEW

    with st.expander(
        "Figure Extraction and Review"
    ):

        figure_review(
            df,
            resolve_image
        )

    # DOWNLOAD DATABASE

    st.subheader(
        "Download Database and Template"
    )

    a, b = st.columns(2)

    a.download_button(
        "Download Admin Database",
        csv_bytes(df),
        CSV.name,
        "text/csv"
    )

    b.download_button(
        "Download Empty CSV Template",
        csv_bytes(
            pd.DataFrame(columns=df.columns)
        ),
        "TEDDIE_Import_Template.csv",
        "text/csv"
    )

    # IMPORT ANOTHER EXAM

    st.subheader("Import Another Exam")

    upload = st.file_uploader(
        "Upload a formatted exam CSV",
        type=["csv"]
    )

    if upload:

        try:

            incoming = normalize(
                read_csv(
                    BytesIO(upload.getvalue())
                )
            )

            st.caption(
                "Review and edit metadata and "
                "content before downloading. "
                "IDs are preserved."
            )

            edited = st.data_editor(
                incoming,
                disabled=[ID, GROUP],
                hide_index=True,
                key=(
                    "import_"
                    + hashlib.sha256(
                        upload.getvalue()
                    ).hexdigest()[:16]
                )
            )

            incoming = normalize(edited)

            errors = []

            if incoming.empty:

                errors.append("Empty import.")

            for col in REQUIRED:

                if (incoming[col] == "").any():

                    errors.append(
                        f"Missing values in {col}"
                    )

            if not incoming[BLOOM].isin(
                LEVELS
            ).all():

                errors.append(
                    "Invalid Bloom's levels."
                )

            if (
                incoming[ID].duplicated().any()
                or incoming[ID].isin(df[ID]).any()
            ):

                errors.append(
                    "Duplicate Question IDs."
                )

            combined = pd.concat(
                [df, incoming],
                ignore_index=True
            ).fillna("")

            if combined.duplicated(
                [EXAM, PARENT, QUESTION],
                keep=False
            ).iloc[len(df):].any():

                errors.append(
                    "Potential duplicate questions."
                )

            for error in errors:

                st.error(error)

            if not errors:

                _, report = missing_report(
                    incoming
                )

                if not report.empty:

                    st.warning(
                        "Add these figures before "
                        "publishing."
                    )

                    st.dataframe(
                        report,
                        hide_index=True
                    )

                st.write(
                    f"Existing: {len(df)} | "
                    f"New: {len(incoming)} | "
                    f"Candidate total: {len(combined)}"
                )

                if st.checkbox(
                    "I reviewed the questions, answers, "
                    "classifications and figure "
                    "associations against the "
                    "original exam."
                ):

                    incoming[
                        "Review Status"
                    ] = "Reviewed"

                    combined = pd.concat(
                        [df, incoming],
                        ignore_index=True
                    ).fillna("")

                    st.download_button(
                        "Download Reviewed Individual Exam",
                        csv_bytes(incoming),
                        "reviewed_exam.csv",
                        "text/csv"
                    )

                    st.download_button(
                        "Download Proposed Database Update",
                        csv_bytes(combined),
                        "exam_questions_proposed.csv",
                        "text/csv",
                        type="primary"
                    )

                st.info(
                    "Downloads do not change the live "
                    "database. Archive the individual "
                    "CSV, back up the master, and have "
                    "the administrator commit the "
                    "reviewed candidate and figures "
                    "to GitHub."
                )

        except Exception as e:

            st.error(
                f"Import failed: {e}"
            )


# ============================================================
# 14. HELP PAGE
# ============================================================

else:

    st.header("Using TEDDIE")

    st.markdown("""
    ### Build an exam

    Explore using the course, topic, Bloom's
    and question type filters.

    Generate exact question-part counts,
    review the basket, and export Word documents.

    ### Add another exam

    1. Download the CSV template.
    2. Convert the exam using a university-approved
       AI extraction service.
    3. Verify the questions, answers and classifications.
    4. Upload the CSV under Manage Database.
    5. Upload the student PDF under Figure Extraction
       and Review.
    6. Review and download the approved figures.
    7. Publish the candidate database and approved
       figures through the private repository.

    Preserve original exams and individual reviewed
    CSVs in restricted departmental storage.

    ### Security

    Keep the application and repository private.

    No Unity account is required to use the
    deployed application.

    ### Bloom's Taxonomy
    """)

    definitions = [
        "Recall facts",
        "Explain concepts",
        "Use a procedure",
        "Examine relationships",
        "Justify judgments",
        "Design something new"
    ]

    for level, definition in zip(
        LEVELS,
        definitions
    ):

        st.write(
            f"**{level}:** {definition}"
        )
