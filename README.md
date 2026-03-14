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

### Basic Example

```bash
python code.py --input_file input_words.txt
```

**Outputs created natively in the current directory:**
* `untitled.apkg` (The Anki Package)
* `database.xlsx` (The Excel database containing an `input_words` tab)

### Appending to an Existing Deck

If `untitled.apkg` already exists, the script will prompt you:

```text
File path untitled.apkg Exists! should i append it to the current file or create a new file? [a/c]: 
```
* Type **`a`** to append the new notes as a sub-deck inside the existing `.apkg` file. The new words will also be appended to `database.xlsx` as a new sheet.
* Type **`c`** to bypass the existing file and create a brand new file (e.g., `untitled_1.apkg`).

### Command Line Arguments

| Argument | Description | Default |
| :--- | :--- | :--- |
| `--input_file` | **(Required)** Path to the text file containing tab-separated words. | _None_ |
| `--output_file` | The name of the generated Anki package. | `untitled.apkg` |
| `--output_dir` | The directory where the Anki package and Excel database will be saved. | `.` (Current directory) |
| `--excel_db` | The name of the Excel database file. | `database.xlsx` |
