# Timeline of Computing

This project contains the source material and scripts for generating a two-sided
Timeline-style card game about the history of computing. The deck contains 65
cards, from the birth of arithmetic to generative AI. It can be printed in A6, or A7 format.
It is designed to help players, especially pupils and students, reflect on the
major events that shaped the history of computing. The short explanatory text on
the back of each card provides context and supports discussion after each event
has been placed on the timeline.

<img src="cs-timeline_illustration.png" alt="Illustration of CS Timeline game" width="75%">

Game instructions are available in [French](Rules_FR.md) and
[English](Rules_EN.md).

## Requirements

Install Python 3, then install the required packages:

```bash
python3 -m pip install --user python-pptx Pillow
```

## Project Files

| Path | Purpose |
| --- | --- |
| `template.pptx` | Two-sided PowerPoint template: front, then back. |
| `text_card_fr.csv` | French card data. |
| `text_card_en.csv` | English card data. |
| `images/` | Images referenced by the CSV `image` column. |
| `generate_cards.py` | Generates a complete front/back PowerPoint deck. |
| `card_fr.pdf` | Ready-to-print French card deck exported from PowerPoint. |
| `card_en.pdf` | Ready-to-print English card deck exported from PowerPoint. |

## AI-Generated Image Disclosure

The card illustrations were generated with AI image-generation tools, including
Google Gemini, OpenAI ChatGPT, and xAI Grok. They are illustrative assets created
for this educational project and may not be historically exact representations of
the people, objects, or events shown.

## Generate a Card Deck

Run the generator from this directory.

### French Deck

```bash
python3 generate_cards.py \
  --template template.pptx \
  --csv text_card_fr.csv \
  --images images \
  --output card_fr.pptx
```

### English Deck

```bash
python3 generate_cards.py \
  --template template.pptx \
  --csv text_card_en.csv \
  --images images \
  --output card_en.pptx
```

Each command produces 130 slides: 65 card fronts followed by their corresponding
backs. The script validates the required template objects and checks that every
image referenced by the CSV exists before saving the presentation.

## Printable PDF Decks

The repository includes ready-to-print PDF exports generated with Microsoft
PowerPoint:

| File | Language | Contents |
| --- | --- | --- |
| `card_fr.pdf` | French | 65 card fronts and 65 matching backs. |
| `card_en.pdf` | English | 65 card fronts and 65 matching backs. |

To create a new PDF after editing card data, first generate the corresponding
PowerPoint deck. Open the `.pptx` file in PowerPoint, then select **File > Export
> Create PDF/XPS** (or **File > Save As > PDF**). For physical cards, print the
PDF double-sided, using the option that flips on the short edge, then cut the
pages to the desired A6 or A7 card format.

### Generator Arguments

| Argument | Default | Description |
| --- | --- | --- |
| `--template` | `template.pptx` | PowerPoint template containing the two card layouts. |
| `--csv` | `text_card_fr.csv` | Card data source. |
| `--images` | `images` | Folder containing card images. |
| `--output` | `card_fr.pptx` | Generated PowerPoint file. |

## CSV Format

The CSV files must contain the following columns:

| Column | Description |
| --- | --- |
| `date` | Date shown on the card back. |
| `title` | Card title displayed on front and back. |
| `subtitle` | Optional front-side subtitle. Leave empty when unused. |
| `image` | Image filename, relative to `images/`. |
| `text` | Explanatory text displayed on the card back. |

Use the literal sequence `\n` inside `title`, `subtitle`, or `text` to insert
a manual line break in PowerPoint. For example:

```csv
1965,"Moore's Law:\nChips Become More Powerful",,card_1965.png,"..."
```

When `subtitle` is empty, the generator removes its paragraph from the front
layout so the title is vertically centered.

## PowerPoint Template Requirements

`template.pptx` must contain exactly two template slides:

1. **Front slide**: `IMG_FOND`, `IMG_OVERLAY`, and `TITRE`.
   `TITRE` must contain the placeholder text `TITLE` and `SUB-TITLE`.
2. **Back slide**: `IMG_FOND`, `IMG_OVERLAY`, `CONTENU`, and `DATE`.
   `CONTENU` must contain `TITLE` and `CONTENT`; `DATE` must contain `DATE`.

The script replaces only these placeholders, preserving the template's layout,
images, and text formatting. Long back-side text is reduced automatically from
10 pt down to a minimum of 8 pt when needed to fit the content box.

## Contributions

Contributions are welcome. You can open an issue to report an error, suggest a
new card or translation, improve the historical content, or propose an
enhancement to the scripts and documentation. Pull requests are welcome too.
