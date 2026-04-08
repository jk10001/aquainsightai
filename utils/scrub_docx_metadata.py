from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# Namespaces used inside DOCX property XML files
NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "cus": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

CORE_TARGETS = [
    "{%s}creator" % NS["dc"],  # Author
    "{%s}lastModifiedBy" % NS["cp"],  # Last Modified By
    "{%s}identifier" % NS["dc"],  # Sometimes internal IDs
]

CORE_TIME_TARGETS = [
    "{%s}created" % NS["dcterms"],
    "{%s}modified" % NS["dcterms"],
]

APP_TARGETS = [
    "{%s}Company" % NS["ep"],
    "{%s}Manager" % NS["ep"],
    "{%s}HyperlinkBase" % NS["ep"],
]


def _blank_or_remove(root: ET.Element, tags: list[str], *, remove: bool) -> None:
    """
    Either blank tag text or remove the element entirely.
    For DOCX props, blanking is usually safest (keeps schema structure).
    """
    for parent in root.iter():
        for child in list(parent):
            if child.tag in tags:
                if remove:
                    parent.remove(child)
                else:
                    child.text = ""


def _scrub_core_xml(xml_bytes: bytes, *, remove_timestamps: bool) -> bytes:
    root = ET.fromstring(xml_bytes)
    # blank, not remove, to stay compatible
    _blank_or_remove(root, CORE_TARGETS, remove=False)
    if remove_timestamps:
        # timestamps are safe to blank too
        _blank_or_remove(root, CORE_TIME_TARGETS, remove=False)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _scrub_app_xml(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    _blank_or_remove(root, APP_TARGETS, remove=False)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _empty_custom_xml(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    # Remove all <property> children
    for child in list(root):
        root.remove(child)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def scrub_docx_metadata(
    in_docx: str | Path,
    out_docx: str | Path | None = None,
    *,
    overwrite: bool = False,
    remove_timestamps: bool = False,
    remove_custom_properties_entirely: bool = True,
) -> Path:
    """
    Scrub DOCX metadata:
    - Rewrites the DOCX ZIP so there are no duplicate members.
    - Blanks author/editor/company-like fields.
    - Optionally removes timestamps.
    - Optionally deletes docProps/custom.xml (or empties it).
    """
    in_path = Path(in_docx)

    if overwrite:
        if out_docx is not None:
            raise ValueError("Provide either overwrite=True OR out_docx, not both.")
        out_path = in_path
        tmp_out = in_path.with_name(in_path.name + ".scrubbed.tmp.docx")
    else:
        out_path = (
            Path(out_docx)
            if out_docx is not None
            else in_path.with_name(in_path.stem + "_scrubbed.docx")
        )
        if in_path.resolve() == out_path.resolve():
            raise ValueError(
                "Input and output paths are the same. Use overwrite=True for in-place scrubbing."
            )
        tmp_out = out_path

    # Read original and write a fresh zip (prevents duplicate members)
    with zipfile.ZipFile(in_path, "r") as zin, zipfile.ZipFile(tmp_out, "w") as zout:
        for item in zin.infolist():
            name = item.filename

            # Optionally delete custom props entirely
            if remove_custom_properties_entirely and name == "docProps/custom.xml":
                continue

            data = zin.read(name)

            # Scrub core/app/custom where applicable
            if name == "docProps/core.xml":
                try:
                    data = _scrub_core_xml(data, remove_timestamps=remove_timestamps)
                except Exception:
                    # If parsing fails, leave it unchanged rather than corrupting docx
                    pass
            elif name == "docProps/app.xml":
                try:
                    data = _scrub_app_xml(data)
                except Exception:
                    pass
            elif (
                not remove_custom_properties_entirely
            ) and name == "docProps/custom.xml":
                try:
                    data = _empty_custom_xml(data)
                except Exception:
                    pass

            # Preserve original compression settings where possible
            zi = zipfile.ZipInfo(filename=item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zi.internal_attr = item.internal_attr
            zi.flag_bits = item.flag_bits

            # NOTE: Some fields (like extra) are ignored by zipfile on write; that's fine for DOCX.
            zout.writestr(zi, data)

    if overwrite:
        tmp_out.replace(in_path)
        return in_path

    return out_path


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(
        description="Scrub DOCX metadata (Word/LibreOffice-safe)."
    )
    p.add_argument("in_docx", help="Input .docx")
    p.add_argument(
        "--out", default=None, help="Output .docx (default: *_scrubbed.docx)"
    )
    p.add_argument(
        "--overwrite", action="store_true", help="Overwrite the input file in-place."
    )
    p.add_argument(
        "--remove-timestamps",
        action="store_true",
        help="Blank created/modified timestamps too.",
    )
    p.add_argument(
        "--keep-custom-props",
        action="store_true",
        help="Keep custom.xml but empty it (default deletes it).",
    )
    args = p.parse_args()

    out = scrub_docx_metadata(
        args.in_docx,
        args.out,
        overwrite=args.overwrite,
        remove_timestamps=args.remove_timestamps,
        remove_custom_properties_entirely=not args.keep_custom_props,
    )
    print(f"Scrubbed DOCX written to: {out}")
