import os
import sys
import argparse
from resume import extract_resume_data


SUPPORTED_EXTS = {".pdf", ".docx", ".doc", ".png", ".jpg", ".jpeg",
                  ".tiff", ".tif", ".bmp", ".heif"}


def process_file(path: str, excel_path: str) -> None:
    if not os.path.isfile(path):
        print(f"✗ Not a file: {path}")
        return

    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        print(f"✗ Unsupported file type: {path}")
        return

    with open(path, "rb") as f:
        content = f.read()

    try:
        row = extract_resume_data(content, os.path.basename(path), excel_path)
        name = f"{row.get('first_name') or ''} {row.get('last_name') or ''}".strip() or "(no name)"
        print(f"✓ {path} -> {name}")
    except Exception as e:
        print(f"✗ Failed {path}: {e}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract resume data from PDF, DOCX, DOC, or image files and append to an Excel file."
    )
    parser.add_argument(
        "files", nargs="+",
        help="One or more resume files (PDF, DOCX, DOC, PNG, JPG, etc.)"
    )
    parser.add_argument(
        "-o", "--output", default="resumes.xlsx",
        help="Path to the Excel output file (default: resumes.xlsx)"
    )
    args = parser.parse_args()

    print(f"Output file: {args.output}\n")
    for path in args.files:
        process_file(path, args.output)

    print(f"\nDone. Data saved to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())