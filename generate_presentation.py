#!/usr/bin/env python3
"""Create a presentation from a template using data from a Word document."""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from pptx import Presentation


@dataclass
class QuestionItem:
    number: int | None
    theme: str
    question: str
    answer: str = ""
    comment: str = ""
    source: str = ""


FIELD_PATTERNS = {
    "theme": re.compile(r"^\s*(?:\d+\.\s*)?Тематика\s*:\s*(.*)$", re.IGNORECASE),
    "question": re.compile(r"^\s*Вопрос\s*:\s*(.*)$", re.IGNORECASE),
    "answer": re.compile(r"^\s*Ответ\s*:\s*(.*)$", re.IGNORECASE),
    "comment": re.compile(r"^\s*Комментарий\s*:\s*(.*)$", re.IGNORECASE),
    "source": re.compile(r"^\s*Источник\s*:\s*(.*)$", re.IGNORECASE),
}
NUMBER_PATTERN = re.compile(r"^\s*(\d+)\.\s*")


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def parse_questions_from_docx(docx_path: Path) -> list[QuestionItem]:
    """Parse questions from a .docx with blocks containing Тематика/Вопрос/... fields."""
    doc = Document(str(docx_path))
    lines = [p.text.strip() for p in doc.paragraphs]

    items: list[QuestionItem] = []
    current: dict[str, str | int | None] | None = None
    current_field: str | None = None

    def flush_current() -> None:
        nonlocal current, current_field
        if not current:
            return
        theme = normalize_spaces(str(current.get("theme", "")))
        question = normalize_spaces(str(current.get("question", "")))
        if theme and question:
            items.append(
                QuestionItem(
                    number=current.get("number"),
                    theme=theme,
                    question=question,
                    answer=normalize_spaces(str(current.get("answer", ""))),
                    comment=normalize_spaces(str(current.get("comment", ""))),
                    source=normalize_spaces(str(current.get("source", ""))),
                )
            )
        current = None
        current_field = None

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        number_match = NUMBER_PATTERN.match(line)
        if number_match and FIELD_PATTERNS["theme"].match(line):
            flush_current()
            current = {"number": int(number_match.group(1))}

        if current is None:
            current = {"number": None}

        matched_field = None
        for field_name, pattern in FIELD_PATTERNS.items():
            match = pattern.match(line)
            if match:
                matched_field = field_name
                value = match.group(1).strip()
                if value:
                    current[field_name] = value
                elif field_name not in current:
                    current[field_name] = ""
                current_field = field_name
                break

        if matched_field:
            continue

        if current_field:
            existing = str(current.get(current_field, "")).strip()
            current[current_field] = f"{existing} {line}".strip()

    flush_current()
    return items


def replace_placeholder(text: str, placeholder: str, value: str) -> str:
    return re.sub(re.escape(placeholder), value, text, flags=re.IGNORECASE)


def replace_in_text_frame(text_frame, replacements: dict[str, str]) -> None:
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            new_text = run.text
            for placeholder, value in replacements.items():
                new_text = replace_placeholder(new_text, placeholder, value)
            run.text = new_text



def replace_in_shape(shape, replacements: dict[str, str]) -> None:
    if getattr(shape, "has_text_frame", False):
        replace_in_text_frame(shape.text_frame, replacements)

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                replace_in_text_frame(cell.text_frame, replacements)

    if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
        for subshape in shape.shapes:
            replace_in_shape(subshape, replacements)


def fill_slide_placeholders(
    presentation_path: Path,
    output_path: Path,
    slide_number: int,
    replacements: dict[str, str],
) -> None:
    prs = Presentation(str(presentation_path))

    slide_index = slide_number - 1
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise ValueError(
            f"В презентации {len(prs.slides)} слайдов, а запрошен слайд {slide_number}."
        )

    slide = prs.slides[slide_index]
    for shape in slide.shapes:
        replace_in_shape(shape, replacements)

    prs.save(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Берёт первый вопрос из Word.docx и подставляет его тематику и текст вопроса "
            "в 6-й слайд шаблона Presentation1.pptx"
        )
    )
    parser.add_argument("--word", default="Word.docx", type=Path, help="Путь к Word-файлу")
    parser.add_argument(
        "--template",
        default="Presentation1.pptx",
        type=Path,
        help="Путь к шаблону презентации",
    )
    parser.add_argument(
        "--output",
        default="Presentation1_filled.pptx",
        type=Path,
        help="Путь для сохранения новой презентации",
    )
    args = parser.parse_args()

    questions = parse_questions_from_docx(args.word)
    if not questions:
        raise ValueError("В Word-файле не найдено ни одного корректного блока с Тематикой и Вопросом.")

    first = questions[0]
    fill_slide_placeholders(
        presentation_path=args.template,
        output_path=args.output,
        slide_number=6,
        replacements={
            "тематика": first.theme,
            "вопрос": first.question,
        },
    )

    print(f"Готово: создан файл {args.output}")
    print(f"Подставлено: тематика='{first.theme}', вопрос='{first.question}'")


if __name__ == "__main__":
    main()
