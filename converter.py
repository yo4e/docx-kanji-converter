"""Core conversion logic for Japanese fiction manuscripts in DOCX files."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.document import Document as DocumentObject
from docx.shared import Pt

RULE_INDENT = "indent"
RULE_NUMBERS = "numbers"
RULE_ELLIPSIS = "ellipsis"
RULE_ASCII = "ascii"
RULE_PUNCTUATION = "punctuation"
RULE_ITALIC = "italic"
RULE_FONT = "font"
RULE_REPLACEMENTS = "replacements"

ALL_RULES = (
    RULE_INDENT,
    RULE_NUMBERS,
    RULE_ELLIPSIS,
    RULE_ASCII,
    RULE_PUNCTUATION,
    RULE_ITALIC,
    RULE_FONT,
    RULE_REPLACEMENTS,
)

_NUMBER_PATTERN = re.compile(r"[0-9]+")
_PUNCTUATION_PATTERN = re.compile(r"([！？])(?![　）」』])")
_HEADING_PREFIXES = ("見出し", "Heading")
_NO_INDENT_PREFIXES = ("　", "「", "（", "『")
_ASCII_FULLWIDTH_TRANSLATION = str.maketrans(
    {
        **{chr(code): chr(code + 0xFEE0) for code in range(ord("A"), ord("Z") + 1)},
        **{chr(code): chr(code + 0xFEE0) for code in range(ord("a"), ord("z") + 1)},
    }
)


@dataclass(frozen=True)
class LiteralReplacement:
    """One exact text replacement applied before the built-in text rules."""

    source: str
    target: str

    def __post_init__(self) -> None:
        if not self.source:
            raise ValueError("replacement source must not be empty")


DEFAULT_LITERAL_REPLACEMENTS = (
    LiteralReplacement("%", "パーセント"),
    LiteralReplacement("％", "パーセント"),
)


def merge_literal_replacements(
    *groups: Iterable[LiteralReplacement],
) -> tuple[LiteralReplacement, ...]:
    """Merge replacement groups, letting later entries override a source."""

    merged: dict[str, str] = {}
    for group in groups:
        for replacement in group:
            merged[replacement.source] = replacement.target
    return tuple(
        LiteralReplacement(source, target) for source, target in merged.items()
    )


def load_literal_replacements(path: Path) -> tuple[LiteralReplacement, ...]:
    """Load UTF-8 tab-separated replacement rules.

    Empty lines and lines beginning with ``#`` are ignored. Each remaining line
    must contain ``source<TAB>target``. An empty target is allowed so a source can
    be deleted, but an empty source is rejected.
    """

    replacements: list[LiteralReplacement] = []
    text = path.read_text(encoding="utf-8-sig")
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        if "\t" not in raw_line:
            raise ValueError(
                f"{path}:{line_number}: expected source and target separated by a tab"
            )
        source, target = raw_line.split("\t", 1)
        if not source:
            raise ValueError(f"{path}:{line_number}: replacement source is empty")
        replacements.append(LiteralReplacement(source, target))
    return tuple(replacements)


@dataclass(frozen=True)
class ConversionOptions:
    """Rules enabled for one conversion run."""

    enabled_rules: frozenset[str] = frozenset(ALL_RULES)
    literal_replacements: tuple[LiteralReplacement, ...] = (
        DEFAULT_LITERAL_REPLACEMENTS
    )

    @classmethod
    def with_disabled(
        cls,
        disabled_rules: Iterable[str],
        *,
        literal_replacements: Iterable[LiteralReplacement] | None = None,
    ) -> "ConversionOptions":
        disabled = frozenset(disabled_rules)
        unknown = disabled.difference(ALL_RULES)
        if unknown:
            names = ", ".join(sorted(unknown))
            raise ValueError(f"Unknown conversion rule(s): {names}")
        replacements = (
            DEFAULT_LITERAL_REPLACEMENTS
            if literal_replacements is None
            else tuple(literal_replacements)
        )
        return cls(
            frozenset(ALL_RULES).difference(disabled),
            tuple(replacements),
        )

    def enabled(self, rule: str) -> bool:
        return rule in self.enabled_rules


@dataclass
class ConversionReport:
    """Counts of applied and skipped changes."""

    changes: Counter[str] = field(default_factory=Counter)
    skipped: Counter[str] = field(default_factory=Counter)

    @property
    def total_changes(self) -> int:
        return sum(self.changes.values())

    def format_text(self) -> str:
        lines = [f"Total changes: {self.total_changes}"]
        for rule in ALL_RULES:
            lines.append(f"- {rule}: {self.changes[rule]}")
        if self.skipped:
            lines.append("Skipped:")
            for reason, count in sorted(self.skipped.items()):
                lines.append(f"- {reason}: {count}")
        return "\n".join(lines)


def convert_number_to_kanji(num_str: str) -> str:
    """Convert a one-to-four digit ASCII number to Japanese numerals.

    Numbers longer than four digits are intentionally rejected. The caller can
    then leave them unchanged instead of risking a crash or incorrect output.
    """

    if not re.fullmatch(r"[0-9]{1,4}", num_str):
        raise ValueError("num_str must contain one to four ASCII digits")

    number = int(num_str)
    if number == 0:
        return "零"

    units = ("", "十", "百", "千")
    digits = ("", "一", "二", "三", "四", "五", "六", "七", "八", "九")
    result = ""

    for position, char in enumerate(reversed(num_str)):
        digit = int(char)
        if digit == 0:
            continue
        if position > 0 and digit == 1:
            result = units[position] + result
        else:
            result = digits[digit] + units[position] + result

    return result


def convert_numbers_in_text(text: str) -> str:
    """Convert ASCII number sequences of at most four digits in text."""

    def replace(match: re.Match[str]) -> str:
        value = match.group(0)
        if len(value) > 4:
            return value
        return convert_number_to_kanji(value)

    return _NUMBER_PATTERN.sub(replace, text)


def replace_ellipsis(text: str) -> str:
    """Replace three ASCII periods with a paired Japanese ellipsis."""

    return text.replace("...", "……")


def convert_ascii_to_fullwidth(text: str) -> str:
    """Convert ASCII Latin letters to full-width Latin letters."""

    return text.translate(_ASCII_FULLWIDTH_TRANSLATION)


def insert_space_after_punctuation(text: str) -> str:
    """Insert a full-width space after Japanese !/? when appropriate."""

    return _PUNCTUATION_PATTERN.sub(r"\1　", text)


def is_heading(style_name: str) -> bool:
    return style_name.startswith(_HEADING_PREFIXES)


def is_heading_paragraph(paragraph) -> bool:
    """Detect headings even when an import has flattened paragraph styles.

    Apple Pages and some office converters export titles and section headings
    with the ``Normal`` paragraph style while retaining bold, larger runs. Treat
    a paragraph as a heading when it has an explicit heading style, or when all
    non-empty runs are bold and at least one is 13pt or larger.
    """

    style_name = paragraph.style.name if paragraph.style else ""
    if is_heading(style_name):
        return True

    content_runs = [run for run in paragraph.runs if run.text.strip()]
    if not content_runs or not all(run.bold is True for run in content_runs):
        return False

    explicit_sizes = [
        run.font.size.pt for run in content_runs if run.font.size is not None
    ]
    return bool(explicit_sizes) and max(explicit_sizes) >= 13


def should_indent(
    style_name: str, text: str, *, heading_like: bool = False
) -> bool:
    """Return whether a paragraph should receive one full-width indent."""

    if not text or not text.strip():
        return False
    if heading_like or is_heading(style_name):
        return False
    return not text.startswith(_NO_INDENT_PREFIXES)


def _apply_text_rules(
    text: str, options: ConversionOptions, report: ConversionReport
) -> str:
    if options.enabled(RULE_REPLACEMENTS):
        for replacement in options.literal_replacements:
            if replacement.source == replacement.target:
                continue
            count = text.count(replacement.source)
            if count:
                text = text.replace(replacement.source, replacement.target)
                report.changes[RULE_REPLACEMENTS] += count

    if options.enabled(RULE_NUMBERS):
        skipped_large = sum(
            1 for match in _NUMBER_PATTERN.finditer(text) if len(match.group(0)) > 4
        )
        if skipped_large:
            report.skipped["numbers_over_4_digits"] += skipped_large
        report.changes[RULE_NUMBERS] += sum(
            1
            for match in _NUMBER_PATTERN.finditer(text)
            if len(match.group(0)) <= 4
            and convert_number_to_kanji(match.group(0)) != match.group(0)
        )
        text = convert_numbers_in_text(text)

    if options.enabled(RULE_ELLIPSIS):
        count = text.count("...")
        if count:
            text = replace_ellipsis(text)
            report.changes[RULE_ELLIPSIS] += count

    if options.enabled(RULE_ASCII):
        changed_chars = sum(
            1 for char in text if ("A" <= char <= "Z") or ("a" <= char <= "z")
        )
        if changed_chars:
            text = convert_ascii_to_fullwidth(text)
            report.changes[RULE_ASCII] += changed_chars

    if options.enabled(RULE_PUNCTUATION):
        text, count = _PUNCTUATION_PATTERN.subn(r"\1　", text)
        report.changes[RULE_PUNCTUATION] += count

    return text


def process_document(
    document: DocumentObject, options: ConversionOptions | None = None
) -> ConversionReport:
    """Apply enabled rules to body paragraphs in an in-memory DOCX document."""

    options = options or ConversionOptions()
    report = ConversionReport()

    for paragraph in document.paragraphs:
        style_name = paragraph.style.name if paragraph.style else ""
        heading_like = is_heading_paragraph(paragraph)

        if (
            options.enabled(RULE_INDENT)
            and paragraph.runs
            and should_indent(
                style_name, paragraph.text, heading_like=heading_like
            )
        ):
            paragraph.runs[0].text = "　" + paragraph.runs[0].text
            report.changes[RULE_INDENT] += 1

        for run in paragraph.runs:
            run.text = _apply_text_rules(run.text, options, report)

            if options.enabled(RULE_ITALIC) and run.italic is True:
                run.italic = False
                run.bold = True
                report.changes[RULE_ITALIC] += 1

            if (
                options.enabled(RULE_FONT)
                and not heading_like
                and run.font.name is not None
            ):
                run.font.name = None
                report.changes[RULE_FONT] += 1

            if (
                options.enabled(RULE_FONT)
                and not heading_like
                and run.font.size != Pt(12)
            ):
                run.font.size = Pt(12)
                report.changes[RULE_FONT] += 1

    return report


def convert_document(
    input_path: Path,
    output_path: Path | None,
    *,
    options: ConversionOptions | None = None,
    dry_run: bool = False,
) -> ConversionReport:
    """Load, convert, and optionally save a DOCX file."""

    document = Document(str(input_path))
    report = process_document(document, options)

    if not dry_run:
        if output_path is None:
            raise ValueError("output_path is required unless dry_run is enabled")
        document.save(str(output_path))

    return report
