# Anki Deck & Excel Database Generator

A Python script that reads a text file containing tab-separated word pairs (e.g., a word and its translation) and generates an Anki flashcard deck (`.apkg`). It also automatically saves these words into a master Excel database (`.xlsx`), creating a new tab for each input file.

## Features

* **Anki Deck Generation**: Creates an Anki `BasicModel` deck with Front/Back cards.
* **Deck Appending**: If the output `.apkg` file already exists, it can append the new words as a sub-deck to the existing Anki package, avoiding the need to import multiple packages manually.
* **Excel Vocabulary Database**: Automatically backs up all processed words into an Excel workbook (`database.xlsx`).
* **Tabbed Organization**: Creates a new Excel tab (sheet) named after the input file for every run.

## Requirements

* Python 3.x
* `genanki`
* `pandas`
* `openpyxl`

You can install the dependencies using pip or conda:

```bash
pip install genanki pandas openpyxl
# OR
conda install pandas openpyxl && pip install genanki
```

## Input File Format

The input must be a plain text file (`.txt`) where each line contains a Front and Back value separated by a **tab** character.

**Example (`input_words.txt`):**
```text
hello	bonjour
world	monde
cat	chat
```

## Usage

Run the script via the command line, providing the path to your input text file.

### Command Line Arguments

| Argument | Description | Default |
| :--- | :--- | :--- |
| `--input_file` | **(Required)** Path to the text file containing tab-separated flashcards. | _None_ |
| `--output_file` | Name of the generated Anki package (`.apkg`). | `deck.apkg` |
| `--output_dir` | Directory where the Anki package and Excel database will be saved. | `.` (current directory) |
| `--excel_db` | Name of the Excel database file used to store imported flashcards. | `database.xlsx` |
| `--deck_name` | Name of the deck inside Anki. Use `::` to create nested decks (e.g., `German::Food`). If omitted, the deck name defaults to `German::<input_file_name>`. | `German::<input_file_name>` |

### Example

```bash
python create_deck.py \
    --input_file vocabulary.txt \
    --deck_name "German::B1::Food" \
    --output_file german_b1.apkg
```

This creates an Anki deck with the following hierarchy:

```
German
└── B1
    └── Food
```

### Appending to an Existing Deck

If `untitled.apkg` already exists, the script will prompt you:

```text
File path untitled.apkg Exists! should i append it to the current file or create a new file? [a/c]: 
```
* Type **`a`** to append the new notes as a sub-deck inside the existing `.apkg` file. The new words will also be appended to `database.xlsx` as a new sheet.
* Type **`c`** to bypass the existing file and create a brand new file (e.g., `untitled_1.apkg`).
