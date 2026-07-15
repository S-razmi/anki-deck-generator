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
# Formatting
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
# Append to existing APKG
# -------------------------
def append_to_apkg(
    existing_apkg: str,
    new_deck: genanki.Deck,
    output_apkg: str | None = None,
) -> None:
    if output_apkg is None:
        output_apkg = existing_apkg

    temp_dir = tempfile.mkdtemp()

    try:
        with zipfile.ZipFile(existing_apkg, "r") as archive:
            archive.extractall(temp_dir)

        db_path = os.path.join(temp_dir, "collection.anki2")

        conn = sqlite3.connect(db_path)

        try:
            cursor = conn.cursor()

            timestamp = time.time()
            id_gen = itertools.count(int(timestamp * 1000))

            new_deck.write_to_db(cursor, timestamp, id_gen)

            conn.commit()
        finally:
            conn.close()

        with zipfile.ZipFile(output_apkg, "w", zipfile.ZIP_DEFLATED) as outzip:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    outzip.write(file_path, arcname)

    finally:
        shutil.rmtree(temp_dir)


# -------------------------
# Excel logging
# -------------------------
def append_to_excel(
    excel_path: Path,
    sheet_name: str,
    data: list[tuple[str, str]],
) -> None:
    df = pd.DataFrame(data, columns=["Front", "Back"])

    safe_sheet_name = sheet_name[:31]

    if not excel_path.exists():
        df.to_excel(
            excel_path,
            sheet_name=safe_sheet_name,
            index=False,
        )
    else:
        with pd.ExcelWriter(
            excel_path,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="new",
        ) as writer:
            df.to_excel(
                writer,
                sheet_name=safe_sheet_name,
                index=False,
            )


# -------------------------
# Main
# -------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create an Anki deck from a tab-separated text file."
    )

    parser.add_argument(
        "--input_file",
        type=str,
        required=True,
        help="Path to the tab-separated input file.",
    )

    parser.add_argument(
        "--output_file",
        type=str,
        default="deck.apkg",
        help="Name of the generated APKG file.",
    )

    parser.add_argument(
        "--output_dir",
        type=str,
        default=".",
        help="Directory where the APKG and Excel files will be saved.",
    )

    parser.add_argument(
        "--excel_db",
        type=str,
        default="database.xlsx",
        help="Name of the Excel database file.",
    )

    parser.add_argument(
        "--deck_name",
        type=str,
        default=None,
        help=(
            "Name of the deck inside Anki. "
            "Use :: to create subdecks, for example German::Food."
        ),
    )

    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    input_path = Path(args.input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if not input_path.is_file():
        raise ValueError(f"Input path is not a file: {input_path}")

    sheet_name = input_path.stem

    # Use the provided deck name.
    # Otherwise, derive it from the input filename.
    deck_name = args.deck_name or f"German::{sheet_name}"

    # Generate a deck ID.
    deck_id = random.randrange(1 << 30, 1 << 31)

    # -------------------------
    # Anki note model
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
    words_data: list[tuple[str, str]] = []

    with input_path.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            line = line.rstrip("\n")

            if not line.strip():
                continue

            if "\t" not in line:
                raise ValueError(
                    f"Line {line_number} is not tab-separated: {line!r}"
                )

            front, back = line.split("\t", maxsplit=1)

            front = front.strip()
            back = back.strip()

            if not front or not back:
                raise ValueError(
                    f"Line {line_number} has an empty field: {line!r}"
                )

            words_data.append((front, back))

            formatted_back = format_answer(back)

            note = genanki.Note(
                model=model,
                fields=[
                    front,
                    formatted_back,
                ],
            )

            deck.add_note(note)

    if not words_data:
        raise ValueError("The input file contains no valid flashcards.")

    # -------------------------
    # Output handling
    # -------------------------
    output_path = out_dir / args.output_file

    if output_path.exists():
        answer = input(
            f"{output_path} already exists. "
            "Append to it or create a new file? [a/c]: "
        ).strip().lower()

        if answer == "a":
            append_to_apkg(
                existing_apkg=str(output_path),
                new_deck=deck,
                output_apkg=str(output_path),
            )

            print(f"Appended deck '{deck_name}' to: {output_path}")

        else:
            output_path = next_available(output_path)

            genanki.Package(deck).write_to_file(str(output_path))

            print(f"Created deck '{deck_name}': {output_path}")

    else:
        genanki.Package(deck).write_to_file(str(output_path))

        print(f"Created deck '{deck_name}': {output_path}")

    # -------------------------
    # Save to Excel
    # -------------------------
    excel_path = out_dir / args.excel_db

    append_to_excel(
        excel_path=excel_path,
        sheet_name=sheet_name,
        data=words_data,
    )

    print(
        f"Saved to Excel: {excel_path} "
        f"(sheet: {sheet_name[:31]})"
    )


if __name__ == "__main__":
    main()
