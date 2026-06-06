#!/usr/bin/env python3
"""
Convert a Word (.docx) file with multiple-choice questions into JSON.

Expected Word structure:
- Theme headings like: Tema 1_Nom del tema
- Questions can be numbered ("1. Pregunta...") or unnumbered if Word numbering hides the number
- Options can be explicit ("a) Opció") or Word auto-lettered lists
- Correct option marked in bold
- Previous exam questions optionally marked in italic

Usage:
    python convert_word_to_json.py test_formated.docx
    python convert_word_to_json.py test_formated.docx --out preguntes.json --js preguntes.js --report informe_validacio.txt
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document

# Theme headings like "Tema 1_Name", "Tema17_Name" or "Tema 50.1_Name".
THEME_RE = re.compile(r"^\s*Tema\s*(\d+(?:\.\d+)?)\s*[_\-–:]\s*(.+?)\s*$", re.IGNORECASE)
# A question number must be at the beginning and followed by a dot/parenthesis.
# The dot cannot be followed by another digit; otherwise amounts such as
# "500.000€" or wrapped years like "2002)." can be misread as questions.
# This still accepts both "1. Text" and "1.Text".
QUESTION_RE = re.compile(r"^\s*(\d{1,3})\s*(?:\.(?!\d)|\))\s*(.+?)\s*$")
OPTION_RE = re.compile(r"^\s*([a-d])\s*[\.)]\s*(.+?)\s*$")
EMBEDDED_OPTION_RE = re.compile(r"(?<!\w)([a-d])\s*[\.)]\s+")
OPTION_LABELS = ["a", "b", "c", "d"]


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


@dataclass
class LogicalLine:
    text: str
    bold: bool
    italic: bool
    paragraph_index: int
    is_list_item: bool = False


@dataclass
class OptionDraft:
    text: str
    is_correct: bool = False
    is_italic: bool = False


@dataclass
class QuestionDraft:
    theme_number: int | None
    theme_name: str
    number: int
    text: str
    question_italic: bool = False
    options: list[OptionDraft] = field(default_factory=list)

    def convocatoria_anterior(self) -> bool:
        if self.question_italic:
            return True
        if not self.options:
            return False
        return sum(1 for opt in self.options if opt.is_italic) >= max(1, len(self.options) // 2)

    def to_json(self) -> dict[str, Any]:
        correct_indices = [i for i, opt in enumerate(self.options) if opt.is_correct]
        correct_index = correct_indices[0] if len(correct_indices) == 1 else None
        return {
            "id": make_question_id(self.theme_number, self.number, self.text),
            "tema_num": self.theme_number,
            "tema": self.theme_name,
            "numero_original": self.number,
            "pregunta": clean_text(self.text),
            "opcions": [clean_text(opt.text) for opt in self.options],
            "correcta": correct_index,
            "correcta_lletra": OPTION_LABELS[correct_index] if correct_index is not None and correct_index < 4 else None,
            "convocatoria_anterior": self.convocatoria_anterior(),
            "explicacio": "",
        }


def make_question_id(theme_number: int | None, question_number: int, question: str) -> str:
    base = f"tema{theme_number or 'x'}-{question_number}-{question}"
    digest = hashlib.sha1(base.encode("utf-8")).hexdigest()[:8]
    return f"T{theme_number or 'X'}_Q{question_number}_{digest}"


def paragraph_to_lines(para, paragraph_index: int) -> list[LogicalLine]:
    """Split a Word paragraph into logical lines, preserving bold/italic per line.

    This is important because some Word paragraphs contain embedded line breaks:
    question + a) + b) + c) + d) all inside the same paragraph.
    """
    ppr = getattr(para._p, "pPr", None)
    is_list_item = ppr is not None and ppr.numPr is not None
    lines: list[dict[str, Any]] = [{"parts": [], "bold": False, "italic": False}]

    for run in para.runs:
        chunks = run.text.split("\n")
        for idx, chunk in enumerate(chunks):
            if idx > 0:
                lines.append({"parts": [], "bold": False, "italic": False})
            if chunk:
                lines[-1]["parts"].append(chunk)
                if run.bold is True and chunk.strip():
                    lines[-1]["bold"] = True
                if run.italic is True and chunk.strip():
                    lines[-1]["italic"] = True

    result: list[LogicalLine] = []
    for line in lines:
        text = clean_text("".join(line["parts"]))
        if text:
            result.append(
                LogicalLine(
                    text=text,
                    bold=line["bold"],
                    italic=line["italic"],
                    paragraph_index=paragraph_index,
                    is_list_item=is_list_item,
                )
            )
    return result


def is_numbered_question(text: str) -> bool:
    return bool(QUESTION_RE.match(text)) and not bool(OPTION_RE.match(text))


def split_embedded_options(text: str) -> list[str]:
    """Split one line containing several explicit options, e.g.
    'a) text b) text c) text d) text', into separate option strings.
    """
    matches = list(EMBEDDED_OPTION_RE.finditer(text))
    if len(matches) <= 1:
        return [text]

    parts: list[str] = []
    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        part = text[start:end].strip()
        if part:
            parts.append(part)
    return parts

def is_probable_unnumbered_question(text: str, line: LogicalLine) -> bool:
    if OPTION_RE.match(text):
        return False

    if line.bold and text.rstrip().endswith(("?", ":", "...", "…")):
        return True

    if text.rstrip().endswith("?") and len(text.split()) >= 4:
        return True

    QUESTION_STARTERS = (
        "marca la resposta",
        "quina ",
        "quin ",
        "quins ",
        "quines ",
        "segons ",
        "d’acord ",
        "d'acord ",
        "en relació ",
        "pel que fa ",
        "sobre ",
    )

    if line.bold and text.lower().startswith(QUESTION_STARTERS):
        return True

    return False


def parse_docx(path: Path) -> tuple[list[QuestionDraft], list[str]]:
    doc = Document(path)
    questions: list[QuestionDraft] = []
    warnings: list[str] = []

    current_theme_number: int | None = None
    current_theme_name = "Sense tema"
    current_question_counter = 0
    current: QuestionDraft | None = None

    def flush_current() -> None:
        nonlocal current
        if current is not None:
            questions.append(current)
            current = None

    for p_index, para in enumerate(doc.paragraphs, start=1):
        for line in paragraph_to_lines(para, p_index):
            text = line.text

            theme_match = THEME_RE.match(text)
            if theme_match:
                flush_current()
                raw_theme_number = theme_match.group(1)
                current_theme_number = int(raw_theme_number.split(".")[0])
                current_theme_name = clean_text(theme_match.group(2))
                if "." in raw_theme_number:
                    current_theme_name = f"{raw_theme_number}_{current_theme_name}"
                current_question_counter = 0
                continue

            opt_match = OPTION_RE.match(text)

            # Numbered question.
            if is_numbered_question(text):
                q_match = QUESTION_RE.match(text)
                assert q_match is not None
                q_number = int(q_match.group(1))
                q_text = clean_text(q_match.group(2))

                # Wrapped continuation such as: "13.On es va acordar..." followed by another line before options.
                if current is not None and len(current.options) == 0 and not current.text.rstrip().endswith(("?", ":", "...", "…")):
                    current.text = clean_text(current.text + " " + text)
                    current.question_italic = current.question_italic or line.italic
                    continue

                flush_current()
                current_question_counter = max(current_question_counter, q_number)
                current = QuestionDraft(
                    theme_number=current_theme_number,
                    theme_name=current_theme_name,
                    number=q_number,
                    text=q_text,
                    question_italic=line.italic,
                )
                continue

            # Explicit option, e.g. a) Text. Also handles accidental same-line
            # options, e.g. 'a) Text b) Text'.
            if opt_match and current is not None:
                for option_text in split_embedded_options(text):
                    split_match = OPTION_RE.match(option_text)
                    if split_match is None:
                        continue
                    current.options.append(
                        OptionDraft(text=clean_text(split_match.group(2)), is_correct=line.bold, is_italic=line.italic)
                    )
                continue
            # If we already have 4 options and this looks like a new unnumbered question,
            # start a new question instead of appending it to option d).
            if current is not None and len(current.options) >= 4 and is_probable_unnumbered_question(text, line):
                flush_current()
                current_question_counter += 1
                current = QuestionDraft(
                    theme_number=current_theme_number,
                    theme_name=current_theme_name,
                    number=current_question_counter,
                    text=text,
                    question_italic=line.italic,
                )
                continue

            # If we already have four options and the next line is not a new
            # question, it is probably a wrapped continuation of option d).
            if current is not None and len(current.options) >= 4 and not is_probable_unnumbered_question(text, line):
                current.options[-1].text = clean_text(current.options[-1].text + " " + text)
                current.options[-1].is_correct = current.options[-1].is_correct or line.bold
                current.options[-1].is_italic = current.options[-1].is_italic or line.italic
                warnings.append(f"Text afegit a l'opció anterior, paràgraf {line.paragraph_index}: {text}")
                continue

            # Start unnumbered question: common when Word hides the numbering.
            if current is None or len(current.options) >= 4:
                if is_probable_unnumbered_question(text, line):
                    flush_current()
                    current_question_counter += 1
                    current = QuestionDraft(
                        theme_number=current_theme_number,
                        theme_name=current_theme_name,
                        number=current_question_counter,
                        text=text,
                        question_italic=line.italic,
                    )
                    continue
                else:
                    warnings.append(f"Línia ignorada abans de cap pregunta, paràgraf {line.paragraph_index}: {text}")
                    continue

            # Wrapped question text before options.
            if current is not None and len(current.options) == 0 and not current.text.rstrip().endswith(("?", ":", "...", "…")):
                current.text = clean_text(current.text + " " + text)
                current.question_italic = current.question_italic or line.italic
                continue

            # Word auto-lettered option where a)/b)/c)/d) is not visible to
            # python-docx. We only treat it as a new option if Word says this
            # paragraph is a list item. Otherwise, it is a wrapped continuation
            # of the previous option. This prevents amounts like '500.000€' or
            # wrapped text like '(Cimera de Johannesburg 2002).' from becoming
            # fake questions/options.
            if current is not None and len(current.options) < 4 and line.is_list_item:
                current.options.append(OptionDraft(text=text, is_correct=line.bold, is_italic=line.italic))
                continue

            # Continuation text after one or more options.
            if current is not None and len(current.options) > 0:
                current.options[-1].text = clean_text(current.options[-1].text + " " + text)
                current.options[-1].is_correct = current.options[-1].is_correct or line.bold
                current.options[-1].is_italic = current.options[-1].is_italic or line.italic
                warnings.append(f"Text afegit a l'opció anterior, paràgraf {line.paragraph_index}: {text}")
                continue

            # Last fallback: possible unlabelled option in a plain paragraph.
            if current is not None and len(current.options) < 4:
                current.options.append(OptionDraft(text=text, is_correct=line.bold, is_italic=line.italic))
                continue

    flush_current()
    return questions, warnings


def validate_questions(questions: list[QuestionDraft], parser_warnings: list[str]) -> str:
    lines: list[str] = []
    lines.append("INFORME DE VALIDACIÓ")
    lines.append("====================")
    lines.append(f"Preguntes totals: {len(questions)}")

    by_theme: dict[str, int] = {}
    previous_count = 0
    errors: list[str] = []
    warnings: list[str] = list(parser_warnings)
    seen: dict[str, list[str]] = {}

    for q in questions:
        theme_label = f"Tema {q.theme_number}: {q.theme_name}" if q.theme_number is not None else q.theme_name
        by_theme[theme_label] = by_theme.get(theme_label, 0) + 1
        if q.convocatoria_anterior():
            previous_count += 1

        location = f"Tema {q.theme_number}, pregunta {q.number}"
        if len(q.options) != 4:
            errors.append(f"{location}: té {len(q.options)} opcions, no 4.")

        correct_indices = [i for i, opt in enumerate(q.options) if opt.is_correct]
        if len(correct_indices) == 0:
            errors.append(f"{location}: cap opció marcada en negreta com a correcta.")
        elif len(correct_indices) > 1:
            letters = ", ".join(OPTION_LABELS[i] for i in correct_indices if i < len(OPTION_LABELS))
            errors.append(f"{location}: més d'una opció marcada en negreta ({letters}).")

        seen.setdefault(clean_text(q.text).lower(), []).append(location)

    duplicates = {k: v for k, v in seen.items() if len(v) > 1}

    lines.append("")
    lines.append("Preguntes per tema:")
    for theme, count in by_theme.items():
        lines.append(f"- {theme}: {count}")

    lines.append("")
    lines.append(f"Preguntes marcades com a convocatòria anterior/cursiva: {previous_count}")

    lines.append("")
    lines.append("Errors:")
    lines.extend([f"- {e}" for e in errors] if errors else ["- Cap error crític detectat."])

    lines.append("")
    lines.append("Avisos:")
    lines.extend([f"- {w}" for w in warnings] if warnings else ["- Cap avís."])

    lines.append("")
    lines.append("Duplicats possibles:")
    if duplicates:
        for _, locs in duplicates.items():
            lines.append(f"- {'; '.join(locs)}")
    else:
        lines.append("- Cap duplicat detectat.")

    return "\n".join(lines) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert Word test questions to JSON.")
    parser.add_argument("docx", type=Path, help="Input .docx file")
    parser.add_argument("--out", type=Path, default=Path("preguntes.json"), help="Output JSON file")
    parser.add_argument("--report", type=Path, default=Path("informe_validacio.txt"), help="Validation report file")
    parser.add_argument("--js", type=Path, default=Path("preguntes.js"), help="Output JavaScript file for double-click local app")
    args = parser.parse_args()

    questions, warnings = parse_docx(args.docx)
    data = [q.to_json() for q in questions]
    args.out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    args.js.write_text("window.PREGUNTES = " + json.dumps(data, ensure_ascii=False, indent=2) + ";\n", encoding="utf-8")
    report = validate_questions(questions, warnings)
    args.report.write_text(report, encoding="utf-8")

    print(report)
    print(f"JSON escrit a: {args.out}")
    print(f"JavaScript escrit a: {args.js}")
    print(f"Informe escrit a: {args.report}")


if __name__ == "__main__":
    main()
