"""Generate front/back PowerPoint card decks from a template and a CSV file."""

import argparse
import copy
import csv
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Pt


DEFAULT_TEMPLATE = "template.pptx"
DEFAULT_CSV = "text_card_fr.csv"
DEFAULT_IMAGE_DIR = "images"
DEFAULT_OUTPUT = "card_fr.pptx"

CSV_COLUMNS = ("date", "title", "subtitle", "image", "text")
FRONT_SHAPES = ("IMG_FOND", "IMG_OVERLAY", "TITRE")
BACK_SHAPES = ("IMG_FOND", "IMG_OVERLAY", "CONTENU", "DATE")
RELATIONSHIP_NAMESPACE = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
)

CONTENT_MIN_FONT_SIZE_PT = 8
CONTENT_FONT_SIZE_STEP_PT = 0.5
LINE_SPACING_FACTOR = 1.2

_MEASURING_DRAW = ImageDraw.Draw(Image.new("RGB", (1, 1)))
_MEASURING_FONT_CACHE = {}


@dataclass(frozen=True)
class Card:
    """One card row from the source CSV."""

    date: str
    title: str
    subtitle: str
    image: str
    text: str


def normalize_text(value):
    """Convert literal CSV line-break markers to actual line breaks."""

    return value.strip().replace("\\n", "\n")


def read_cards(csv_path):
    """Load and validate cards from a CSV file."""

    csv_path = Path(csv_path)
    if not csv_path.is_file():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    with csv_path.open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        missing_columns = set(CSV_COLUMNS) - set(reader.fieldnames or [])

        if missing_columns:
            raise ValueError(
                f"Missing CSV columns: {', '.join(sorted(missing_columns))}"
            )

        cards = []
        for line_number, row in enumerate(reader, start=2):
            card = Card(
                date=normalize_text(row["date"]),
                title=normalize_text(row["title"]),
                subtitle=normalize_text(row["subtitle"]),
                image=row["image"].strip(),
                text=normalize_text(row["text"]),
            )
            validate_card(card, line_number)
            cards.append(card)

    if not cards:
        raise ValueError("The CSV does not contain any cards.")

    return cards


def validate_card(card, line_number):
    """Ensure required values are present in a card row."""

    for field_name, value in {
        "image": card.image,
        "title": card.title,
        "date": card.date,
        "text": card.text,
    }.items():
        if not value:
            raise ValueError(f"Line {line_number}: empty {field_name} value.")


def find_shape(slide, name):
    """Return the top-level shape with the requested PowerPoint name."""

    for shape in slide.shapes:
        if shape.name == name:
            return shape

    available_shapes = [shape.name for shape in slide.shapes]
    raise ValueError(
        f"Object '{name}' not found on the slide. "
        f"Available objects: {available_shapes}"
    )


def validate_template(presentation):
    """Ensure that the two template slides contain their required shapes."""

    if len(presentation.slides) < 2:
        raise ValueError(
            "The PowerPoint template must contain at least two slides: front and back."
        )

    for side_name, slide, required_shapes in (
        ("front", presentation.slides[0], FRONT_SHAPES),
        ("back", presentation.slides[1], BACK_SHAPES),
    ):
        available_shapes = {shape.name for shape in slide.shapes}
        missing_shapes = set(required_shapes) - available_shapes

        if missing_shapes:
            raise ValueError(
                f"Missing {side_name} template objects: "
                f"{', '.join(sorted(missing_shapes))}. "
                f"Available objects: {', '.join(sorted(available_shapes))}"
            )


def replace_image(shape, image_path):
    """Replace an existing picture while preserving its layout and crop."""

    image_path = Path(image_path)
    if not image_path.is_file():
        raise FileNotFoundError(f"Image not found: {image_path}")

    _, relationship_id = shape.part.get_or_add_image_part(str(image_path))
    shape._element.blipFill.blip.set(
        RELATIONSHIP_NAMESPACE + "embed", relationship_id
    )


def find_placeholder_run(shape, placeholder):
    """Return the run and paragraph whose text exactly matches a placeholder."""

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.text == placeholder:
                return run, paragraph

    raise ValueError(
        f"Placeholder text '{placeholder}' was not found in shape '{shape.name}'."
    )


def set_run_text(run, text):
    """Set text while preserving formatting across PowerPoint soft line breaks."""

    lines = text.replace("\x0b", " ").split("\n")
    run.text = lines[0]
    anchor = run._r

    for line in lines[1:]:
        line_break = anchor.makeelement(qn("a:br"), {})
        run_properties = anchor.find(qn("a:rPr"))
        if run_properties is not None:
            line_break.append(copy.deepcopy(run_properties))

        new_run = copy.deepcopy(anchor)
        new_run.find(qn("a:t")).text = line
        anchor.addnext(line_break)
        line_break.addnext(new_run)
        anchor = new_run


def replace_placeholder_text(shape, placeholder, text):
    """Replace a placeholder without changing its formatting."""

    run, _ = find_placeholder_run(shape, placeholder)
    set_run_text(run, text)


def remove_placeholder_paragraph(shape, placeholder):
    """Remove the paragraph that contains a placeholder."""

    _, paragraph = find_placeholder_run(shape, placeholder)
    paragraph._p.getparent().remove(paragraph._p)


def copy_slide(presentation, source_slide):
    """Duplicate a slide, including shapes and their required relationships."""

    new_slide = presentation.slides.add_slide(source_slide.slide_layout)
    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)

    relationship_map = copy_relationships(source_slide, new_slide)
    for shape in source_slide.shapes:
        copied_shape = copy.deepcopy(shape.element)
        remap_shape_relationships(copied_shape, relationship_map)
        new_slide.shapes._spTree.insert_element_before(copied_shape, "p:extLst")

    return new_slide


def copy_relationships(source_slide, target_slide):
    """Copy non-layout relationships and return a map of old to new IDs."""

    relationship_map = {}
    for relationship_id, relationship in source_slide.part.rels.items():
        if "notesSlide" in relationship.reltype or "slideLayout" in relationship.reltype:
            continue

        target = (
            relationship.target_ref
            if relationship.is_external
            else relationship.target_part
        )
        relationship_map[relationship_id] = target_slide.part.relate_to(
            target,
            relationship.reltype,
            is_external=relationship.is_external,
        )

    return relationship_map


def remap_shape_relationships(shape_element, relationship_map):
    """Replace copied XML relationship IDs with IDs valid on the target slide."""

    for element in shape_element.iter():
        for attribute in ("embed", "link", "id"):
            old_id = element.get(RELATIONSHIP_NAMESPACE + attribute)
            if old_id in relationship_map:
                element.set(
                    RELATIONSHIP_NAMESPACE + attribute,
                    relationship_map[old_id],
                )


def delete_slide(presentation, slide):
    """Delete a slide from a presentation."""

    for slide_id in presentation.slides._sldIdLst:
        if int(slide_id.id) == slide.slide_id:
            presentation.part.drop_rel(slide_id.rId)
            presentation.slides._sldIdLst.remove(slide_id)
            return


def get_measuring_font(size_pt):
    """Return a cached approximate font used only for layout estimation."""

    size_px = max(1, round(size_pt))
    if size_px not in _MEASURING_FONT_CACHE:
        _MEASURING_FONT_CACHE[size_px] = ImageFont.load_default(size=size_px)
    return _MEASURING_FONT_CACHE[size_px]


def count_wrapped_lines(text, font_size_pt, max_width_pt):
    """Estimate display lines after wrapping the supplied text."""

    font = get_measuring_font(font_size_pt)
    line_count = 0

    for segment in text.split("\n"):
        words = segment.split()
        if not words:
            line_count += 1
            continue

        line = ""
        for word in words:
            candidate = f"{line} {word}".strip()
            if line and _MEASURING_DRAW.textlength(candidate, font=font) > max_width_pt:
                line_count += 1
                line = word
            else:
                line = candidate
        line_count += 1

    return line_count


def content_fits(shape, title, text, title_size_pt, text_size_pt):
    """Estimate whether a card back's title and text fit its content box."""

    text_frame = shape.text_frame
    width_pt = (shape.width - text_frame.margin_left - text_frame.margin_right) / 12700
    height_pt = (shape.height - text_frame.margin_top - text_frame.margin_bottom) / 12700
    title_lines = count_wrapped_lines(title, title_size_pt, width_pt)
    text_lines = count_wrapped_lines(text, text_size_pt, width_pt)
    required_height_pt = LINE_SPACING_FACTOR * (
        title_lines * title_size_pt + text_size_pt + text_lines * text_size_pt
    )
    return required_height_pt <= height_pt


def fit_content_font_size(shape, title, text, title_size_pt, base_size_pt):
    """Return the largest content font size that fits the card back."""

    font_size_pt = base_size_pt
    while font_size_pt > CONTENT_MIN_FONT_SIZE_PT:
        if content_fits(shape, title, text, title_size_pt, font_size_pt):
            return font_size_pt
        font_size_pt = round(font_size_pt - CONTENT_FONT_SIZE_STEP_PT, 1)
    return CONTENT_MIN_FONT_SIZE_PT


def configure_front(slide, card, image_path):
    """Fill a front template slide with one card's image and title information."""

    replace_image(find_shape(slide, "IMG_FOND"), image_path)
    title_shape = find_shape(slide, "TITRE")
    replace_placeholder_text(title_shape, "TITLE", card.title)

    if card.subtitle:
        replace_placeholder_text(title_shape, "SUB-TITLE", card.subtitle)
    else:
        remove_placeholder_paragraph(title_shape, "SUB-TITLE")


def configure_back(slide, card, image_path):
    """Fill a back template with a card's image, date, title, and text."""

    replace_image(find_shape(slide, "IMG_FOND"), image_path)
    content_shape = find_shape(slide, "CONTENU")
    title_run, _ = find_placeholder_run(content_shape, "TITLE")
    text_run, text_paragraph = find_placeholder_run(content_shape, "CONTENT")
    title_size_pt = title_run.font.size.pt if title_run.font.size else 14
    base_text_size_pt = text_run.font.size.pt if text_run.font.size else 10

    replace_placeholder_text(content_shape, "TITLE", card.title)
    replace_placeholder_text(content_shape, "CONTENT", card.text)

    text_size_pt = fit_content_font_size(
        content_shape,
        card.title,
        card.text,
        title_size_pt,
        base_text_size_pt,
    )
    for run in text_paragraph.runs:
        run.font.size = Pt(text_size_pt)

    replace_placeholder_text(find_shape(slide, "DATE"), "DATE", card.date)


def generate_cards(template_path, csv_path, image_dir, output_path):
    """Generate a complete front/back deck from a template, CSV, and image folder."""

    template_path = Path(template_path)
    image_dir = Path(image_dir)
    if not template_path.is_file():
        raise FileNotFoundError(f"PowerPoint template not found: {template_path}")

    cards = read_cards(csv_path)
    presentation = Presentation(template_path)
    validate_template(presentation)
    template_slides = list(presentation.slides)
    front_template, back_template = template_slides[0], template_slides[1]

    print(f"Generating {len(cards)} cards...")
    for number, card in enumerate(cards, start=1):
        image_path = image_dir / card.image
        if not image_path.is_file():
            raise FileNotFoundError(f"Image not found for '{card.title}': {image_path}")

        print(f"[{number}/{len(cards)}] {card.title}")
        configure_front(copy_slide(presentation, front_template), card, image_path)
        configure_back(copy_slide(presentation, back_template), card, image_path)

    delete_slide(presentation, front_template)
    delete_slide(presentation, back_template)
    presentation.save(output_path)

    print("\nGeneration complete")
    print(f"Cards generated: {len(cards)}")
    print(f"Slides generated: {len(cards) * 2}")
    print(f"File: {output_path}")


def parse_args():
    """Parse command-line arguments."""

    parser = argparse.ArgumentParser(
        description="Generate PowerPoint cards from a template and a CSV file."
    )
    parser.add_argument("--template", default=DEFAULT_TEMPLATE, help="PowerPoint template file")
    parser.add_argument("--csv", default=DEFAULT_CSV, help="CSV file containing card data")
    parser.add_argument("--images", default=DEFAULT_IMAGE_DIR, help="Folder containing images")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Output PowerPoint file")
    return parser.parse_args()


def main():
    """Run the command-line generator."""

    args = parse_args()
    generate_cards(args.template, args.csv, args.images, args.output)


if __name__ == "__main__":
    main()
