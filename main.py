"""Command-line entry point for docx-kanji-converter."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError

from converter import (
    ALL_RULES,
    DEFAULT_LITERAL_REPLACEMENTS,
    ConversionOptions,
    convert_document,
    load_literal_replacements,
    merge_literal_replacements,
)


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}.converted{input_path.suffix}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Apply Japanese fiction manuscript formatting rules to a DOCX file."
    )
    parser.add_argument("input", type=Path, help="Input .docx file")
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        help="Output .docx file (default: <input>.converted.docx)",
    )
    parser.add_argument(
        "--disable",
        action="append",
        default=[],
        choices=ALL_RULES,
        metavar="RULE",
        help="Disable a conversion rule; may be repeated",
    )
    parser.add_argument(
        "--replacement-file",
        action="append",
        default=[],
        type=Path,
        metavar="PATH",
        help=(
            "Load UTF-8 tab-separated literal replacements; may be repeated. "
            "Later files override earlier entries with the same source."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report changes without writing an output file",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Allow replacing an existing output file",
    )
    return parser


def validate_paths(
    parser: argparse.ArgumentParser,
    input_path: Path,
    output_path: Path,
    *,
    dry_run: bool,
    force: bool,
) -> None:
    if not input_path.is_file():
        parser.error(f"input file does not exist: {input_path}")
    if input_path.suffix.lower() != ".docx":
        parser.error("input file must have a .docx extension")

    if dry_run:
        return

    try:
        same_file = input_path.resolve() == output_path.resolve()
    except OSError:
        same_file = input_path.absolute() == output_path.absolute()
    if same_file:
        parser.error("refusing to overwrite the input file")
    if output_path.exists() and not force:
        parser.error(f"output file already exists: {output_path} (use --force to replace it)")
    if output_path.suffix.lower() != ".docx":
        parser.error("output file must have a .docx extension")
    if not output_path.parent.exists():
        parser.error(f"output directory does not exist: {output_path.parent}")


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path: Path = args.input
    output_path: Path = args.output or default_output_path(input_path)
    validate_paths(
        parser,
        input_path,
        output_path,
        dry_run=args.dry_run,
        force=args.force,
    )

    try:
        replacements = DEFAULT_LITERAL_REPLACEMENTS
        for replacement_path in args.replacement_file:
            replacements = merge_literal_replacements(
                replacements, load_literal_replacements(replacement_path)
            )
        options = ConversionOptions.with_disabled(
            args.disable, literal_replacements=replacements
        )
        report = convert_document(
            input_path,
            None if args.dry_run else output_path,
            options=options,
            dry_run=args.dry_run,
        )
    except (BadZipFile, PackageNotFoundError, OSError, ValueError) as error:
        parser.exit(1, f"error: {error}\n")

    print(report.format_text())
    if args.dry_run:
        print("Dry run: no output file was written.")
    else:
        print(f"Saved: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
