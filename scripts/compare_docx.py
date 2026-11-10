#!/usr/bin/env python3
"""Read-only comparison of two DOCX files, including non-text structure."""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def read_docx(path: Path) -> tuple[dict[str, bytes], list[str], bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise ValueError(f"Corrupt ZIP member in {path}: {bad}")
        parts = {name: zf.read(name) for name in zf.namelist()}
    if "word/document.xml" not in parts:
        raise ValueError(f"word/document.xml missing in {path}")
    root = ET.fromstring(parts["word/document.xml"])
    paragraphs = ["".join(t.text or "" for t in p.findall(".//w:t", NS))
                  for p in root.findall(".//w:body//w:p", NS)]
    return parts, paragraphs, root


def normalize_document(root: ET.Element) -> bytes:
    for t in root.findall(".//w:t", NS):
        t.text = "TEXT"
    return ET.tostring(root, encoding="utf-8")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source", type=Path, help="original template")
    ap.add_argument("output", type=Path, help="new DOCX to verify")
    args = ap.parse_args()
    src, out = args.source.resolve(), args.output.resolve()
    if not src.is_file() or not out.is_file():
        print(json.dumps({"error": "file_not_found", "source_exists": src.is_file(), "output_exists": out.is_file()}, ensure_ascii=False)); return 2
    try:
        a, ap, ar = read_docx(src)
        b, bp, br = read_docx(out)
        count = max(len(ap), len(bp))
        changed = []
        for i in range(count):
            av = ap[i] if i < len(ap) else None
            bv = bp[i] if i < len(bp) else None
            if av != bv:
                changed.append({"paragraph_index_1_based": i + 1, "source": av, "output": bv})
        same_names = set(a) == set(b)
        other_names = sorted((set(a) | set(b)) - {"word/document.xml"})
        other_diffs = [name for name in other_names if a.get(name) != b.get(name)]
        structure_same = normalize_document(ar) == normalize_document(br)
        result = {
            "source": str(src), "source_sha256": sha(src),
            "output": str(out), "output_sha256": sha(out),
            "package_member_names_same": same_names,
            "non_document_parts_identical": not other_diffs,
            "different_non_document_parts": other_diffs,
            "document_xml_structure_same_after_text_masking": structure_same,
            "source_paragraph_count": len(ap), "output_paragraph_count": len(bp),
            "changed_paragraphs": changed,
            "only_document_text_changed": same_names and not other_diffs and structure_same,
            "limitations": ["This checks package and XML structure, not rendered page appearance.",
                            "Text inside headers, footers, comments, footnotes, or embedded objects is reported through package differences only, not paragraph-by-paragraph."],
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["only_document_text_changed"] else 1
    except (ValueError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(json.dumps({"error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False)); return 2


if __name__ == "__main__":
    sys.exit(main())
