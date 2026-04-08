from __future__ import annotations

import argparse
from pathlib import Path

import fitz  # PyMuPDF


def scrub_pdf_metadata(
    in_pdf: str | Path,
    out_pdf: str | Path | None = None,
    *,
    overwrite: bool = False,
    remove_xmp: bool = True,
    remove_embedded_files: bool = True,
) -> Path:
    """
    Scrub PDF metadata using PyMuPDF (fitz).

    - Clears standard document metadata (author/title/subject/keywords/etc.)
    - Optionally removes XMP metadata (XML packet)
    - Optionally removes embedded file attachments
    - Supports safe overwrite (temp file + atomic replace)

    Returns output path.
    """
    in_path = Path(in_pdf)

    if overwrite:
        if out_pdf is not None:
            raise ValueError("Provide either overwrite=True OR out_pdf, not both.")
        out_path = in_path
        tmp_path = in_path.with_name(in_path.name + ".scrubbed.tmp.pdf")
    else:
        out_path = (
            Path(out_pdf)
            if out_pdf is not None
            else in_path.with_name(in_path.stem + "_scrubbed.pdf")
        )
        if in_path.resolve() == out_path.resolve():
            raise ValueError(
                "Input and output paths are the same. Use overwrite=True for in-place scrubbing."
            )
        tmp_path = out_path

    doc = fitz.open(in_path)

    try:
        # 1) Clear standard PDF metadata keys
        # PyMuPDF expects a dict; missing keys are fine.
        doc.set_metadata(
            {
                "title": "",
                "author": "",
                "subject": "",
                "keywords": "",
                "creator": "",
                "producer": "",
                "creationDate": "",
                "modDate": "",
                "trapped": "",
            }
        )

        # 2) Remove XMP metadata if present
        if remove_xmp:
            try:
                doc.set_xml_metadata("")
            except Exception:
                pass

        # 3) Remove embedded file attachments (rare, but possible)
        if remove_embedded_files:
            try:
                # Iteratively delete all embedded files
                while True:
                    names = getattr(doc, "embfile_names", None)
                    if callable(names):
                        emb_names = list(doc.embfile_names())
                        if not emb_names:
                            break
                        for name in emb_names:
                            doc.embfile_del(name)
                    else:
                        break
            except Exception:
                pass

        # Save as a fresh file to ensure changes persist cleanly.
        # garbage=4 cleans unused objects; deflate compresses.
        doc.save(
            tmp_path,
            garbage=4,
            deflate=True,
            incremental=False,
            encryption=fitz.PDF_ENCRYPT_NONE,
        )
    finally:
        doc.close()

    if overwrite:
        tmp_path.replace(in_path)
        return in_path

    return tmp_path


def main() -> int:
    ap = argparse.ArgumentParser(description="Scrub PDF metadata using PyMuPDF (fitz).")
    ap.add_argument("in_pdf", help="Input PDF path")
    ap.add_argument(
        "--out", default=None, help="Output PDF path (default: *_scrubbed.pdf)"
    )
    ap.add_argument(
        "--overwrite", action="store_true", help="Overwrite input PDF in-place"
    )
    ap.add_argument(
        "--keep-xmp", action="store_true", help="Do not remove XMP metadata"
    )
    ap.add_argument(
        "--keep-embedded",
        action="store_true",
        help="Do not remove embedded file attachments",
    )
    args = ap.parse_args()

    out = scrub_pdf_metadata(
        args.in_pdf,
        args.out,
        overwrite=args.overwrite,
        remove_xmp=not args.keep_xmp,
        remove_embedded_files=not args.keep_embedded,
    )
    print(f"Scrubbed PDF written to: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
