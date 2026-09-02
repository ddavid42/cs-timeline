# Timeline of Computing

This folder contains the source material and scripts for generating a two-sided
Timeline-style card game about the history of computing. The deck contains 65
cards, from the birth of arithmetic to generative AI. It can be printed in A6, or A7 format.

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
| `template.pptx` | Two-slide PowerPoint template: front, then back. |
| `text_card_fr.csv` | French card data. |
| `text_card_en.csv` | English card data. |
| `images/` | Images referenced by the CSV `image` column. |
| `generate_cards.py` | Generates a complete front/back PowerPoint deck. |

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
