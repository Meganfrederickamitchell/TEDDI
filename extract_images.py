
"""
TEDDIE Exam Figure Extraction

Extract, review, crop, and publish exam figures.
Never overwrite existing published images.
"""

import argparse
import csv
import re
from pathlib import Path

import pymupdf as fitz
from PIL import Image, ImageOps, ImageDraw


# ============================================================
# SETTINGS
# ============================================================

FIELDS = [
    "PDF Page",
    "Extracted File",
    "Image File",
    "Associated Question"
]


# ============================================================
# VALIDATE FILENAMES
# ============================================================

def safe_png(name):

    if not re.fullmatch(
        r"[A-Za-z0-9_-][A-Za-z0-9_.-]*\.png",
        name,
        re.I
    ):
        raise ValueError(f"Invalid PNG filename: {name}")

    return name


# ============================================================
# RENDER PDF REGIONS
# ============================================================

def render_region(page, rect=None, dpi=360):

    pix = page.get_pixmap(
        dpi=dpi,
        clip=rect,
        alpha=False,
        annots=False
    )

    return Image.frombytes(
        "RGB",
        (pix.width, pix.height),
        pix.samples
    )


# ============================================================
# EXTRACT FIGURES
# ============================================================

def extract(
    pdf,
    exam_id,
    root=Path("extracted_images"),
    dpi=360
):

    if not re.fullmatch(r"[A-Za-z0-9_-]+", exam_id):
        raise ValueError("Invalid exam ID")

    out = Path(root) / exam_id

    records = []

    with fitz.open(pdf) as doc:

        out.mkdir(
            parents=True,
            exist_ok=False
        )

        for n, page in enumerate(doc, 1):

            # Save full-page reference

            reference = render_region(
                page,
                dpi=180
            )

            reference.save(
                out / f"{exam_id}_p{n:02d}_reference.png"
            )

            # Locate embedded image regions

            for i, info in enumerate(
                page.get_image_info(),
                1
            ):

                rect = (
                    fitz.Rect(info["bbox"])
                    & page.rect
                )

                if rect.is_empty or rect.is_infinite:
                    continue

                name = (
                    f"{exam_id}_p{n:02d}_fig{i:02d}.png"
                )

                # Render the visible PDF region
                # rather than extracting raw image bytes

                image = render_region(
                    page,
                    rect,
                    dpi
                )

                image.save(out / name)

                records.append({
                    "PDF Page": n,
                    "Extracted File": name,
                    "Image File": "",
                    "Associated Question": ""
                })

    # Save image manifest

    with (out / "image_manifest.csv").open(
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=FIELDS
        )

        writer.writeheader()
        writer.writerows(records)

    # Generate contact sheets

    for start in range(0, len(records), 12):

        sheet = Image.new(
            "RGB",
            (1200, 1120),
            "white"
        )

        draw = ImageDraw.Draw(sheet)

        for j, row in enumerate(
            records[start:start + 12]
        ):

            x = j % 3 * 400
            y = j // 3 * 280

            with Image.open(
                out / row["Extracted File"]
            ) as im:

                thumb = ImageOps.contain(
                    im,
                    (380, 240)
                )

                sheet.paste(
                    thumb,
                    (x + (400 - thumb.width) // 2, y)
                )

            draw.text(
                (x + 8, y + 245),
                row["Extracted File"],
                fill="black"
            )

        sheet.save(
            out / f"contact_sheet_{start // 12 + 1:02d}.png"
        )

    print(
        f"{len(records)} candidates in {out}; "
        "review before publishing."
    )

    return out


# ============================================================
# MANUAL FIGURE CROPPING
# ============================================================

def crop_figure(
    pdf,
    page_number,
    coords,
    output,
    dpi=360
):

    output = Path(output)

    safe_png(output.name)

    with fitz.open(pdf) as doc:

        if not 1 <= page_number <= len(doc):
            raise ValueError("Invalid page number")

        page = doc[page_number - 1]

        rect = fitz.Rect(coords) & page.rect

        if rect.is_empty:
            raise ValueError("Empty crop")

        # Exclusive creation prevents overwriting

        with output.open("xb") as f:

            render_region(
                page,
                rect,
                dpi
            ).save(
                f,
                format="PNG"
            )

    return output


# ============================================================
# PUBLISH REVIEWED IMAGES
# ============================================================

def publish(folder, destination):

    folder = Path(folder)
    destination = Path(destination)

    with (folder / "image_manifest.csv").open(
        encoding="utf-8-sig",
        newline=""
    ) as f:

        rows = list(csv.DictReader(f))

    plan = {}

    for row in rows:

        target = row["Image File"].strip()

        # Ignore unapproved / unmapped images

        if not target:
            continue

        src = folder / safe_png(
            row["Extracted File"].strip()
        )

        dst = destination / safe_png(target)

        if (
            not src.is_file()
            or dst.exists()
            or dst in plan
        ):
            raise ValueError(
                f"Missing source or duplicate/existing "
                f"destination: {src} -> {dst}"
            )

        with Image.open(src) as im:
            im.verify()

        plan[dst] = src

    if not plan:
        raise ValueError(
            "No reviewed mappings in manifest"
        )

    destination.mkdir(
        parents=True,
        exist_ok=True
    )

    for dst, src in plan.items():

        with dst.open("xb") as f:
            f.write(src.read_bytes())

        print(f"Copied {dst}")


# ============================================================
# COMMAND LINE INTERFACE
# ============================================================

if __name__ == "__main__":

    p = argparse.ArgumentParser(
        description=__doc__
    )

    sub = p.add_subparsers(
        dest="command",
        required=True
    )

    # Extract command

    e = sub.add_parser("extract")

    e.add_argument("pdf")
    e.add_argument("exam_id")

    e.add_argument(
        "--dpi",
        type=int,
        default=360
    )

    # Crop command

    c = sub.add_parser("crop")

    c.add_argument("pdf")
    c.add_argument("page", type=int)
    c.add_argument("coords", type=float, nargs=4)
    c.add_argument("output")

    # Publish command

    s = sub.add_parser("publish")

    s.add_argument("folder")

    s.add_argument(
        "--destination",
        default="IMAGES"
    )

    a = p.parse_args()

    if a.command == "extract":

        extract(
            a.pdf,
            a.exam_id,
            dpi=a.dpi
        )

    elif a.command == "crop":

        crop_figure(
            a.pdf,
            a.page,
            a.coords,
            a.output
        )

    else:

        publish(
            a.folder,
            a.destination
        )
