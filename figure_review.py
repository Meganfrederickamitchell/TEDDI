
"""
TEDDIE Faculty Figure Review

Extract, crop, review and download exam figures.

Does not modify the admin database CSV or published images.
"""

import io
import hashlib
import zipfile

import pymupdf as fitz
import streamlit as st

from extract_images import (
    render_region,
    safe_png
)


# ============================================================
# FIGURE REVIEW INTERFACE
# ============================================================

def figure_review(data, resolve_image):

    st.subheader(
        "Extract and Review Exam Figures"
    )

    st.caption(
        "Use the original student exam PDF whenever "
        "possible. Answer-key figures may contain "
        "completed answers."
    )

    # --------------------------------------------------------
    # 1. UPLOAD PDF
    # --------------------------------------------------------

    upload = st.file_uploader(
        "Exam PDF",
        type=["pdf"],
        key="figure_pdf"
    )

    if upload is None:
        return

    pdf_bytes = upload.getvalue()

    digest = hashlib.sha256(
        pdf_bytes
    ).hexdigest()

    if st.session_state.get(
        "figure_pdf_id"
    ) != digest:

        st.session_state.figure_pdf_id = digest

        st.session_state.figure_stage = {}

    staged = st.session_state.figure_stage

    # --------------------------------------------------------
    # 2. OPEN PDF
    # --------------------------------------------------------

    try:

        doc = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )

    except Exception as e:

        st.error(
            f"Unable to open PDF: {e}"
        )

        return

    with doc:

        if doc.needs_pass:

            st.error(
                "This PDF is password protected."
            )

            return

        if len(doc) == 0:

            st.error("The PDF has no pages.")

            return

        st.success(
            f"PDF loaded: {len(doc)} pages"
        )

        # ----------------------------------------------------
        # 3. SELECT EXAM
        # ----------------------------------------------------

        exam = st.selectbox(
            "Associate with exam",
            sorted(
                data["Exam/Course"].unique()
            ),
            key="figure_exam"
        )

        rows = data[
            data["Exam/Course"] == exam
        ]

        targets = sorted({
            name
            for name in rows["Image File"]
            if name
            and not resolve_image(name)
            and name not in staged
        })

        st.metric(
            "Missing figure filenames",
            len(targets)
        )

        # ----------------------------------------------------
        # 4. SELECT PDF PAGE
        # ----------------------------------------------------

        page_no = st.number_input(
            "PDF page",
            min_value=1,
            max_value=len(doc),
            value=1,
            key=f"figure_page_{digest}"
        )

        page = doc[int(page_no) - 1]

        st.markdown("**Original PDF page**")

        st.image(
            render_region(
                page,
                dpi=120
            ),
            caption=(
                "Page reference only — "
                "not included in the exported figures"
            )
        )

        # ----------------------------------------------------
        # 5. LOCATE AUTOMATIC FIGURES
        # ----------------------------------------------------

        candidates = page.get_image_info()

        st.write(
            f"Embedded image regions detected: "
            f"{len(candidates)}"
        )

        figure_options = [
            "Manual crop"
        ] + [
            f"Figure {i + 1}"
            for i in range(len(candidates))
        ]

        choice = st.selectbox(
            "Figure region",
            figure_options,
            key=(
                f"figure_choice_"
                f"{digest}_{page_no}"
            )
        )

        # ----------------------------------------------------
        # 6. AUTOMATIC OR MANUAL CROP
        # ----------------------------------------------------

        if choice == "Manual crop":

            st.caption(
                "Adjust the horizontal and vertical "
                "boundaries to isolate the complete "
                "figure. Exclude answers and unrelated "
                "question text."
            )

            x = st.slider(
                "Horizontal bounds (% of page)",
                min_value=0,
                max_value=100,
                value=(0, 100),
                key=(
                    f"crop_x_"
                    f"{digest}_{page_no}"
                )
            )

            y = st.slider(
                "Vertical bounds (% of page)",
                min_value=0,
                max_value=100,
                value=(0, 100),
                key=(
                    f"crop_y_"
                    f"{digest}_{page_no}"
                )
            )

            rect = fitz.Rect(

                page.rect.x0
                + x[0] * page.rect.width / 100,

                page.rect.y0
                + y[0] * page.rect.height / 100,

                page.rect.x0
                + x[1] * page.rect.width / 100,

                page.rect.y0
                + y[1] * page.rect.height / 100
            )

        else:

            index = int(
                choice.split()[-1]
            ) - 1

            rect = (
                fitz.Rect(
                    candidates[index]["bbox"]
                )
                & page.rect
            )

        if rect.is_empty or rect.is_infinite:

            st.warning(
                "Select a valid, nonempty crop."
            )

            return

        # ----------------------------------------------------
        # 7. HIGH-RESOLUTION FIGURE PREVIEW
        # ----------------------------------------------------

        st.markdown("**Extracted figure preview**")

        dpi = st.selectbox(
            "Figure resolution",
            [216, 300, 360],
            index=2,
            key="figure_resolution"
        )

        image = render_region(
            page,
            rect,
            dpi=dpi
        )

        st.image(
            image,
            caption=f"Figure preview — {dpi} DPI"
        )

        st.caption(
            f"Image dimensions: "
            f"{image.width} × {image.height} pixels"
        )

        extrema = image.convert(
            "RGB"
        ).getextrema()

        blank = all(
            low == high
            for low, high in extrema
        )

        if blank:

            st.warning(
                "This region appears blank. "
                "Select another figure or crop."
            )

        # ----------------------------------------------------
        # 8. ASSOCIATE FIGURE WITH QUESTIONS
        # ----------------------------------------------------

        st.markdown(
            "**Associate figure with database**"
        )

        target = st.selectbox(
            "Missing image filename",
            [""] + targets,
            key=f"figure_target_{digest}_{exam}"
        )

        name = st.text_input(
            "Destination filename",
            value=target,
            key=(
                f"figure_name_"
                f"{digest}_{exam}_{target}"
            )
        )

        name = name.strip()

        if target:

            associated = rows.loc[
                rows["Image File"] == target,
                [
                    "Part Name",
                    "Question Part Text"
                ]
            ]

            st.write(
                "Questions referencing this figure:"
            )

            st.dataframe(
                associated,
                hide_index=True
            )

        else:

            st.caption(
                "New filenames must also be assigned "
                "to the appropriate questions in the "
                "reviewed CSV."
            )

        # ----------------------------------------------------
        # 9. APPROVE AND STAGE FIGURE
        # ----------------------------------------------------

        approval_key = (
            f"approve_{digest}_{exam}_"
            f"{page_no}_{choice}_{rect}_"
            f"{name}_{dpi}"
        )

        approved = st.checkbox(
            "I verified that this complete figure "
            "is safe for the student exam and "
            "matches the intended questions.",
            key=approval_key
        )

        if st.button(
            "Stage Reviewed Figure",
            type="primary",
            disabled=(
                not approved
                or not name
                or blank
            ),
            key="stage_figure"
        ):

            try:

                safe_png(name)

                if resolve_image(name):

                    raise ValueError(
                        "This filename already exists "
                        "in the published image folder."
                    )

                if name in staged:

                    raise ValueError(
                        "This filename is already staged."
                    )

                buf = io.BytesIO()

                image.save(
                    buf,
                    format="PNG",
                    dpi=(dpi, dpi)
                )

                staged[name] = buf.getvalue()

                st.success(
                    f"Figure approved: {name}"
                )

            except ValueError as e:

                st.error(str(e))

    # --------------------------------------------------------
    # 10. REVIEW STAGED FIGURES
    # --------------------------------------------------------

    st.divider()

    st.subheader("Approved Figures")

    if not staged:

        st.info(
            "No figures have been approved yet."
        )

        return

    st.metric(
        "Reviewed figures staged",
        len(staged)
    )

    for name in list(staged):

        a, b = st.columns([4, 1])

        with a:

            st.write(name)

            st.image(
                staged[name],
                width=250
            )

        with b:

            if st.button(
                "Remove",
                key=f"remove_figure_{name}"
            ):

                del staged[name]

                st.rerun()

    # --------------------------------------------------------
    # 11. DOWNLOAD APPROVED IMAGES
    # --------------------------------------------------------

    buf = io.BytesIO()

    with zipfile.ZipFile(
        buf,
        "w",
        zipfile.ZIP_DEFLATED
    ) as archive:

        for name, content in staged.items():

            archive.writestr(
                "IMAGES/" + name,
                content
            )

    st.download_button(
        "Download Reviewed Figures ZIP",
        data=buf.getvalue(),
        file_name="TEDDIE_reviewed_figures.zip",
        mime="application/zip",
        type="primary"
    )

    st.info(
        "Download your reviewed figures before "
        "leaving this session or uploading another PDF. "
        "An administrator must add the approved "
        "images to the private repository. "
        "No live files or admin database CSV are changed "
        "by this interface."
    )
