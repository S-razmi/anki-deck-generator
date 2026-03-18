import argparse
import itertools
import os
import random
import shutil
import sqlite3
import tempfile
import time
import zipfile
from pathlib import Path

import genanki
import pandas as pd


# -------------------------
# Formatting (FIXED)
# -------------------------
def format_answer(text: str) -> str:
    parts = text.split(maxsplit=1)

    if len(parts) < 2:
        return text

    article, word = parts

    if article in ["der", "die", "das"]:
        return f'<div class="{article}">{article} {word}</div>'

    return text


# -------------------------
# Utility: avoid overwrite
# -------------------------
def next_available(path: Path) -> Path:
    stem = path.stem
    suffix = path.suffix
    i = 1

    while True:
        new_path = path.with_name(f"{stem}_{i}{suffix}")
        if not new_path.exists():
            return new_path
        i += 1


# -------------------------
# Append to existing apkg
# -------------------------
def append_to_apkg(existing_apkg: str, new_deck: genanki.Deck, output_apkg: str = None):
    if output_apkg is None:
        output_apkg = existing_apkg

    temp_dir = tempfile.mkdtemp()

    try:
        with zipfile.ZipFile(existing_apkg, "r") as z:
            z.extractall(temp_dir)

        db_path = os.path.join(temp_dir, "collection.anki2")
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        timestamp = time.time()
        id_gen = itertools.count(int(timestamp * 1000))

        new_deck.write_to_db(cursor, timestamp, id_gen)

        conn.commit()
        conn.close()

        with zipfile.ZipFile(output_apkg, "w") as outzip:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    outzip.write(file_path, arcname)

    finally:
        shutil.rmtree(temp_dir)


# -------------------------
# Excel logging
# -------------------------
def append_to_excel(excel_path: Path, sheet_name: str, data: list):
    df = pd.DataFrame(data, columns=["Front", "Back"])

    if not excel_path.exists():
        df.to_excel(excel_path, sheet_name=sheet_name[:31], index=False)
    else:
        with pd.ExcelWriter(
            excel_path, engine="openpyxl", mode="a", if_sheet_exists="new"
        ) as writer:
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)


# -------------------------
# MAIN
# -------------------------
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_file", type=str, required=True)
    parser.add_argument("--output_file", type=str, default="deck.apkg")
    parser.add_argument("--output_dir", type=str, default=".")
    parser.add_argument("--excel_db", type=str, default="database.xlsx")

    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    input_path = Path(args.input_file)
    sheet_name = input_path.stem

    # -------------------------
    # FIX: dynamic deck naming
    # -------------------------
    deck_name = f"German::{sheet_name}"

    # -------------------------
    # FIX: unique deck ID
    # -------------------------
    deck_id = random.randrange(1 << 30, 1 << 31)

    # -------------------------
    # Model with CSS styling
    # -------------------------
    model = genanki.Model(
        1607392319,
        "GermanArticleModel",
        fields=[
            {"name": "Front"},
            {"name": "Back"},
        ],
        templates=[
            {
                "name": "Card 1",
                "qfmt": "{{Front}}",
                "afmt": '{{Front}}<hr id="answer">{{Back}}',
            },
        ],
        css="""
        .card {
            font-family: Arial;
            font-size: 20px;
            text-align: center;
        }

        .der {
            text-align: left;
            color: blue;
            font-size: 28px;
        }

        .die {
            text-align: right;
            color: red;
            font-size: 28px;
        }

        .das {
            text-align: center;
            color: green;
            font-size: 28px;
        }
        """,
    )

    deck = genanki.Deck(deck_id, deck_name)

    # -------------------------
    # Read input file
    # -------------------------
    words_data = []

    with open(input_path, "r", encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue

            if "\t" not in line:
                raise ValueError(f"Line is not tab-separated: {line}")

            front, back = line.split("\t")
            front = front.strip()
            back = back.strip()

            words_data.append((front, back))

            formatted_back = format_answer(back)

            note = genanki.Note(
                model=model,
                fields=[front, formatted_back],
            )

            deck.add_note(note)

    # -------------------------
    # Output handling
    # -------------------------
    output_path = out_dir / args.output_file

    if output_path.exists():
        answer = input(
            f"{output_path} exists. Append or create new? [a/c]: "
        ).lower()

        if answer == "a":
            append_to_apkg(str(output_path), deck, str(output_path))
            print(f"Appended to {output_path}")
        else:
            output_path = next_available(output_path)
            genanki.Package(deck).write_to_file(output_path)
            print(f"Created {output_path}")
    else:
        genanki.Package(deck).write_to_file(output_path)
        print(f"Created {output_path}")

    # -------------------------
    # Save Excel
    # -------------------------
    excel_path = out_dir / args.excel_db
    append_to_excel(excel_path, sheet_name, words_data)

    print(f"Saved to Excel: {excel_path} (sheet: {sheet_name[:31]})")


if __name__ == "__main__":
    main()
