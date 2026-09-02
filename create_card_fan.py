"""Create a fan visual from five randomly selected card fronts in a PDF deck."""

import argparse
import random
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageFilter

DEFAULT_PDF = "card_en.pdf"
DEFAULT_OUTPUT = "card_en_fan.png"
CARD_COUNT = 5
RENDER_DPI = 180
CANVAS_SIZE = (2400, 1900)
CARD_WIDTH = 500
ZERO_CARD_FRONT_PAGE = 11
ZERO_CARD_BACK_PAGE = 12
ZERO_CARD_WIDTH = 480


def render_page(pdf_path, page_number, output_prefix):
    """Render one PDF page to PNG with Poppler."""

    subprocess.run(
        [
            "pdftoppm",
            "-png",
            "-r",
            str(RENDER_DPI),
            "-f",
            str(page_number),
            "-l",
            str(page_number),
            str(pdf_path),
            str(output_prefix),
        ],
        check=True,
    )

    return next(output_prefix.parent.glob(f"{output_prefix.name}-*.png"))


def composite_card_with_shadow(canvas, card, x, y):
    """Draw a card with an unclipped, evenly distributed soft shadow."""

    shadow_padding = 30
    shadow = Image.new(
        "RGBA",
        (card.width + shadow_padding * 2, card.height + shadow_padding * 2),
        (0, 0, 0, 0),
    )
    shadow.paste(
        (0, 0, 0, 55),
        (shadow_padding, shadow_padding,
         shadow_padding + card.width, shadow_padding + card.height),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(16))

    canvas.alpha_composite(
        shadow,
        (x - shadow_padding, y - shadow_padding),
    )
    canvas.alpha_composite(card, (x, y))


def create_fan(pdf_path, output_path):
    available_front_pages = [
        page_number
        for page_number in range(1, 131, 2)
        if page_number != ZERO_CARD_FRONT_PAGE
    ]
    random_pages = random.sample(available_front_pages, CARD_COUNT)

    with tempfile.TemporaryDirectory() as temporary_directory:
        temporary_path = Path(temporary_directory)
        cards = []

        for page_number in random_pages:
            image_path = render_page(
                pdf_path,
                page_number,
                temporary_path / f"card-{page_number}",
            )
            card = Image.open(image_path).convert("RGBA")
            card.thumbnail((CARD_WIDTH, 10_000), Image.Resampling.LANCZOS)
            cards.append(card)

        zero_front_path = render_page(
            pdf_path,
            ZERO_CARD_FRONT_PAGE,
            temporary_path / "zero-front",
        )
        zero_back_path = render_page(
            pdf_path,
            ZERO_CARD_BACK_PAGE,
            temporary_path / "zero-back",
        )
        zero_cards = [
            Image.open(zero_front_path).convert("RGBA"),
            Image.open(zero_back_path).convert("RGBA"),
        ]

        for card in zero_cards:
            card.thumbnail((ZERO_CARD_WIDTH, 10_000), Image.Resampling.LANCZOS)

        canvas = Image.new("RGBA", CANVAS_SIZE, "white")
        angles = [-16, -8, 0, 8, 16]
        centers = [(485, 690), (820, 590), (1200, 540), (1580, 590), (1915, 690)]

        for card, angle, center in zip(cards, angles, centers):
            rotated_card = card.rotate(
                angle,
                resample=Image.Resampling.BICUBIC,
                expand=True,
            )

            x = center[0] - rotated_card.width // 2
            y = center[1] - rotated_card.height // 2
            composite_card_with_shadow(canvas, rotated_card, x, y)

        for card, center_x in zip(zero_cards, (930, 1470)):
            x = center_x - card.width // 2
            y = 1160
            composite_card_with_shadow(canvas, card, x, y)

        canvas.convert("RGB").save(output_path, quality=95)

    print(f"Selected PDF pages: {', '.join(map(str, random_pages))}")
    print(
        "Included Invention of Zero pages: "
        f"{ZERO_CARD_FRONT_PAGE} (front), {ZERO_CARD_BACK_PAGE} (back)"
    )
    print(f"Visual created: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Create a fan visual from five random card fronts in a PDF deck."
    )
    parser.add_argument("--pdf", default=DEFAULT_PDF, help="Source PDF deck")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Output PNG file")
    args = parser.parse_args()

    create_fan(Path(args.pdf), Path(args.output))


if __name__ == "__main__":
    main()