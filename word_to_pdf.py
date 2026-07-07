"""
word_to_pdf.py — Convert Word (.docx/.doc) to PDF
Uses LibreOffice

Usage:
    python word_to_pdf.py document.docx
    python word_to_pdf.py document.docx --output output_folder/
    python word_to_pdf.py *.docx
    python word_to_pdf.py folder/

Install dependencies:
    Linux/Mac: sudo apt install libreoffice  or  brew install libreoffice
    Windows:   download LibreOffice from https://www.libreoffice.org/
"""

import subprocess
import sys
import os
import glob
from pathlib import Path


def convert_to_pdf(input_path: str, output_dir: str = None) -> bool:
    """Converts a Word file to PDF using LibreOffice."""
    input_path = Path(input_path)

    if not input_path.exists():
        print(f"  [ERROR] File not found: {input_path}")
        return False

    if input_path.suffix.lower() not in [".docx", ".doc", ".odt", ".rtf"]:
        print(f"  [SKIP] Unsupported format: {input_path.suffix}")
        return False

    if output_dir:
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
    else:
        out_dir = input_path.parent

    print(f"  Converting: {input_path.name} → {input_path.stem}.pdf")

    try:
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(out_dir),
                str(input_path)
            ],
            capture_output=True,
            text=True,
            timeout=60
        )

        pdf_path = out_dir / (input_path.stem + ".pdf")

        if result.returncode == 0 and pdf_path.exists():
            size_kb = pdf_path.stat().st_size / 1024
            print(f"  [OK] Saved: {pdf_path} ({size_kb:.1f} KB)")
            return True
        else:
            print(f"  [ERROR] Conversion failed.")
            if result.stderr:
                print(f"  Details: {result.stderr[:200]}")
            return False

    except FileNotFoundError:
        print("  [ERROR] LibreOffice is not installed or not in PATH.")
        print("  Install: sudo apt install libreoffice")
        return False
    except subprocess.TimeoutExpired:
        print("  [ERROR] Timeout — file is too large or LibreOffice is not responding.")
        return False


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    args = sys.argv[1:]
    output_dir = None

    if "--output" in args:
        idx = args.index("--output")
        output_dir = args[idx + 1]
        args = [a for i, a in enumerate(args) if i != idx and i != idx + 1]

    files = []
    for arg in args:
        if os.path.isdir(arg):
            for ext in ["*.docx", "*.doc", "*.odt", "*.rtf"]:
                files.extend(glob.glob(os.path.join(arg, ext)))
        else:
            matches = glob.glob(arg)
            if matches:
                files.extend(matches)
            else:
                files.append(arg)

    if not files:
        print("[ERROR] No files found to convert.")
        sys.exit(1)

    print(f"\n=== Word to PDF Converter ===")
    print(f"Files to convert: {len(files)}")
    if output_dir:
        print(f"Output folder: {output_dir}")
    print()

    ok = 0
    fail = 0

    for f in files:
        success = convert_to_pdf(f, output_dir)
        if success:
            ok += 1
        else:
            fail += 1

    print(f"\n=== Result ===")
    print(f"Successfully converted: {ok}")
    if fail:
        print(f"Failed: {fail}")


if __name__ == "__main__":
    main()
