"""
certsplit.py

Split a multi-page certificate PDF into individual PDFs and name each file
using recipient names extracted from structured certificate text.

This tool works best for repeated certificate templates, such as Canva-generated
certificate PDFs, where the recipient name appears near a known anchor phrase.

Example:
    python certsplit.py certificates.pdf --anchor "The following award is given to" --preview

    python certsplit.py certificates.pdf --anchor "The following award is given to" --write --output "./out" --prompt-on-fail --export-csv report.csv
"""

from __future__ import annotations

import argparse
import csv
import sys
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import fitz  
from pypdf import PdfReader, PdfWriter


@dataclass
class PageResult:
    page_number: int
    detected_name: Optional[str]
    final_name: str
    filename: str
    status: str  # ok | manual | fallback


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00a0", " ")
    return text


def clean_line(line: str) -> str:
    line = normalize_text(line).strip()
    line = re.sub(r"\s+", " ", line)
    return line


def sanitize_filename(name: str, max_len: int = 120) -> str:
    name = unicodedata.normalize("NFKD", name)
    name = "".join(ch for ch in name if not unicodedata.combining(ch))
    name = re.sub(r'[<>:"/\\|?*\n\r\t]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(". ")

    # replace spaces with underscores for cleaner filenames
    name = name.replace(" ", "_")

    if not name:
        name = "unknown_name"

    return name[:max_len]


def extract_lines_from_page(page: fitz.Page) -> list[str]:
    text = page.get_text("text")
    text = normalize_text(text)
    raw_lines = text.splitlines()
    return [clean_line(line) for line in raw_lines if clean_line(line)]


def detect_name_from_certificate(lines: list[str], anchor: str) -> Optional[str]:
    anchor_lower = clean_line(anchor).lower()

    # Phrases we know are definitely NOT the recipient's name
    bad_phrases = {
        "certificate",
        "of participation",
        "the following award is given to",
        "has successfully",
        "daffodil",
        "university",
        "research wing",
        "director",
        anchor_lower
    }

    for idx, line in enumerate(lines):
        current = clean_line(line).lower()

        # Using 'in' is safer than 'startswith' to bypass hidden PDF formatting characters
        if anchor_lower in current:
            
            # Since PDF extraction order is scrambled, check the lines immediately 
            # surrounding the anchor (1 line above, 1 line below, 2 lines above, etc.)
            search_indices = [idx - 1, idx + 1, idx - 2, idx + 2]
            
            for c_idx in search_indices:
                if 0 <= c_idx < len(lines):
                    candidate = clean_line(lines[c_idx])
                    if not candidate:
                        continue
                        
                    candidate_lower = candidate.lower()

                    # 1. Skip lines that contain known template phrases
                    if any(bad in candidate_lower for bad in bad_phrases):
                        continue
                    
                    # 2. Skip random small PDF artifacts like "th" or "nd" (e.g., Line 12)
                    if len(candidate) <= 3:
                        continue
                        
                    # If it passes the checks, we've found our name!
                    return candidate

    return None


def ensure_unique_path(output_dir: Path, base_filename: str, ext: str = ".pdf") -> Path:
    candidate = output_dir / f"{base_filename}{ext}"
    counter = 2

    while candidate.exists():
        candidate = output_dir / f"{base_filename}_{counter}{ext}"
        counter += 1

    return candidate


def make_fallback_name(prefix: str, page_number: int) -> str:
    return f"{prefix}_{page_number:03d}"


def prompt_for_manual_name(page_number: int, fallback_name: str) -> str:
    print(f"\nPage {page_number}: Could not detect recipient name.")
    user_input = input(
        f'Enter a name manually, or press Enter to keep "{fallback_name}": '
    ).strip()

    return user_input if user_input else fallback_name


def preview_results(results: list[PageResult]) -> None:
    print("\nPreview:\n")
    for result in results:
        label = result.detected_name if result.detected_name else f"FAILED -> {result.final_name}"
        print(f"[{result.page_number}] {label}")


def export_csv_report(results: list[PageResult], csv_path: Path) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["page", "detected_name", "final_name", "filename", "status"])
        for result in results:
            writer.writerow([
                result.page_number,
                result.detected_name or "",
                result.final_name,
                result.filename,
                result.status,
            ])


def write_individual_pdfs(
    input_pdf: Path,
    output_dir: Path,
    results: list[PageResult],
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(input_pdf))

    for result in results:
        writer = PdfWriter()
        writer.add_page(reader.pages[result.page_number - 1])

        output_path = ensure_unique_path(output_dir, result.filename, ".pdf")
        with open(output_path, "wb") as f:
            writer.write(f)

        print(f"Saved: {output_path}")


def process_pdf(
    input_pdf: Path,
    anchor: str,
    fallback_prefix: str,
    prompt_on_fail: bool,
) -> list[PageResult]:
    doc = fitz.open(input_pdf)
    results: list[PageResult] = []

    seen_names: dict[str, int] = {}

    for i in range(len(doc)):
        page_number = i + 1
        lines = extract_lines_from_page(doc[i])
        
        detected_name = detect_name_from_certificate(lines, anchor) 

        status = "ok"
        if detected_name:
            final_name = detected_name
        else:
            fallback_name = make_fallback_name(fallback_prefix, page_number)
            if prompt_on_fail:
                manual_name = prompt_for_manual_name(page_number, fallback_name)
                final_name = manual_name
                status = "manual" if manual_name != fallback_name else "fallback"
            else:
                final_name = fallback_name
                status = "fallback"

        safe_filename = sanitize_filename(final_name)

        if safe_filename in seen_names:
            seen_names[safe_filename] += 1
            safe_filename = f"{safe_filename}_{seen_names[safe_filename]}"
        else:
            seen_names[safe_filename] = 1

        results.append(
            PageResult(
                page_number=page_number,
                detected_name=detected_name,
                final_name=final_name,
                filename=safe_filename,
                status=status,
            )
        )

    return results


def print_summary(results: list[PageResult], output_dir: Optional[Path]) -> None:
    total = len(results)
    auto_named = sum(1 for r in results if r.status == "ok")
    manually_corrected = sum(1 for r in results if r.status == "manual")
    fallback_used = sum(1 for r in results if r.status == "fallback")

    print("\nSummary:")
    print(f"Processed: {total} pages")
    print(f"Named automatically: {auto_named}")
    print(f"Manually corrected: {manually_corrected}")
    print(f"Fallback names used: {fallback_used}")
    if output_dir:
        print(f"Output directory: {output_dir.resolve()}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Split a multi-page certificate PDF into individual PDFs named after recipients."
    )

    parser.add_argument("input_pdf", help="Path to the input multi-page PDF")
    parser.add_argument(
        "--anchor",
        required=True,
        help="Anchor text that appears immediately before the recipient name",
    )
    parser.add_argument(
        "--output",
        default="output",
        help="Directory to save individual PDFs (default: output)",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Preview detected names without writing files",
    )
    parser.add_argument(
        "--write",
        action="store_true",
        help="Write individual PDF files",
    )
    parser.add_argument(
        "--prompt-on-fail",
        action="store_true",
        help="Prompt for manual name entry when extraction fails",
    )
    parser.add_argument(
        "--export-csv",
        default=None,
        help="Optional path to save a CSV processing report",
    )
    parser.add_argument(
        "--fallback-prefix",
        default="page",
        help='Prefix for fallback filenames (default: "page")',
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    input_pdf = Path(args.input_pdf)
    output_dir = Path(args.output)

    if not input_pdf.exists():
        print(f"Error: file not found: {input_pdf}", file=sys.stderr)
        return 1

    if not args.preview and not args.write:
        print("Error: choose at least one action: --preview or --write", file=sys.stderr)
        return 1

    try:
        results = process_pdf(
            input_pdf=input_pdf,
            anchor=args.anchor,
            fallback_prefix=args.fallback_prefix,
            prompt_on_fail=args.prompt_on_fail,
        )

        if args.preview:
            preview_results(results)

        if args.export_csv:
            export_csv_report(results, Path(args.export_csv))
            print(f"\nCSV report saved to: {Path(args.export_csv).resolve()}")

        if args.write:
            write_individual_pdfs(input_pdf, output_dir, results)

        print_summary(results, output_dir if args.write else None)
        return 0

    except KeyboardInterrupt:
        print("\nOperation cancelled by user.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())