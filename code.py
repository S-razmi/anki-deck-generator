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

def format_answer(text: str) -> str:
    parts = text.split(maxsplit=1)

    if len(parts) < 2:
        return text

    article, word = parts

    if article == "der":
        align = "left"
        color = "blue"
    elif article == "die":
        align = "right"
        color = "red"
    elif article == "das":
        align = "center"
        color = "green"
    else:
        return text

    return f'<div style="text-align:{align}; font-size:28px;"><span style="color:{color}">{article}</span> {word}</div>'

def next_available(path: Path) -> Path:
    stem = path.stem  # "file"
    suffix = path.suffix  # ".txt"
    i = 1

    while True:
        new_path = path.with_name(f"{stem}_{i}{suffix}")
        if not new_path.exists():
            return new_path
        i += 1


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


def append_to_excel(excel_path: Path, sheet_name: str, data: list):
    df = pd.DataFrame(data, columns=["Front", "Back"])
    if not excel_path.exists():
        df.to_excel(excel_path, sheet_name=sheet_name[:31], index=False)
    else:
        with pd.ExcelWriter(
            excel_path, engine="openpyxl", mode="a", if_sheet_exists="new"
        ) as writer:
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)


model = genanki.Model(
    1607392319,
    "BasicModel",
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
)

parser = argparse.ArgumentParser()
parser.add_argument("--input_file", type=str, required=True)
parser.add_argument("--output_file", type=str, default="untitled.apkg")
parser.add_argument("--output_dir", type=str, default=".")
parser.add_argument("--excel_db", type=str, default="database.xlsx")

args = parser.parse_args()

out_dir = Path(args.output_dir)
out_dir.mkdir(parents=True, exist_ok=True)

input_path = Path(args.input_file)
sheet_name = input_path.stem

words_data = []
with open(input_path, "r", encoding="utf-8") as f:
    for line in f:
        if not line.strip():
            continue
        front, back = line.split("\t")  # tab separated
        words_data.append((front.strip(), back.strip()))

output_path = out_dir / args.output_file

append_mode = False
deck_name = "Generated Deck"
deck_id = 2059400222

if output_path.exists():
    answer = input(
        f"File path {output_path} Exists! should i append it to the current file or create a new file? [a/c]: "
    )
    if answer == "" or answer.lower() == "c":
        output_path = next_available(output_path)
    elif answer.lower() == "a":
        append_mode = True
        deck_name = f"Generated Deck::{sheet_name}"
        deck_id = random.randrange(1 << 30, 1 << 31)
    else:
        print("Please answer with a or c.")
        exit(1)

deck = genanki.Deck(deck_id, deck_name)
for front, back in words_data:
    formatted_back = format_answer(back)
    note = genanki.Note(model=model, fields=[front, formatted_back])
    deck.add_note(note)

if append_mode:
    append_to_apkg(str(output_path), deck, str(output_path))
    print(f"{output_path} successfully appended.")
else:
    genanki.Package(deck).write_to_file(output_path)
    print(f"{output_path} successfully created.")

excel_path = out_dir / args.excel_db
append_to_excel(excel_path, sheet_name, words_data)
print(f"Words successfully saved to {excel_path} in tab '{sheet_name[:31]}'.")
# type: ignore
