#!/usr/bin/env python3
"""Safely replace unique text contained in individual DOCX w:t nodes.

Read/write only the main document text; preserves all other OOXML members and
run formatting. Refuses ambiguous, split-run, missing, or in-place edits.
Standard library only.
"""
from __future__ import annotations

import argparse
import hashlib
import html
import io
import json
import os
import re
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
WT = f"{{{W}}}t"
TEXT_RE = re.compile(rb"(<w:t\b[^>]*>)(.*?)(</w:t\s*>)", re.DOTALL)
XML_SPACE = b'xml:space="preserve"'


def parse_pairs(items: list[str]) -> list[tuple[str, str]]:
    pairs = []
    for item in items:
        if "=>" not in item:
            raise ValueError(f"Replacement must use OLD=>NEW: {item}")
        old, new = item.split("=>", 1)
        if not old:
            raise ValueError("Old text cannot be empty")
        pairs.append((old, new))
    if len({old for old, _ in pairs}) != len(pairs):
        raise ValueError("Duplicate old-text values are not allowed")
    return pairs


def apply_to_xml(xml: bytes, pairs: list[tuple[str, str]], replace_all: bool) -> tuple[bytes, list[dict]]:
    # Validate package XML first; byte-preserving substitutions below avoid
    # rewriting namespace prefixes, mc:Ignorable values, or unrelated XML.
    root = ET.fromstring(xml)
    t_nodes = root.findall(".//w:t", NS)
    if not t_nodes:
        raise ValueError("No Word text nodes (w:t) found")
    for match in TEXT_RE.finditer(xml):
        # Verify this common prefix-based patcher covers every Word text node.
        pass
    matches = list(TEXT_RE.finditer(xml))
    if len(matches) != len(t_nodes):
        raise ValueError("Unsupported DOCX namespace prefix or text-node form; no output written")

    report = []
    spans_by_node = [[] for _ in matches]
    for old, new in pairs:
        counts = []
        for node_index, m in enumerate(matches):
            decoded = html.unescape(m.group(2).decode("utf-8"))
            positions = []
            start = 0
            while True:
                pos = decoded.find(old, start)
                if pos < 0:
                    break
                positions.append(pos)
                spans_by_node[node_index].append((pos, pos + len(old)))
                start = pos + 1
            counts.append(len(positions))
        count = sum(counts)
        paragraph_count = 0
        for paragraph in root.findall('.//w:p', NS):
            joined = ''.join((t.text or '') for t in paragraph.findall('.//w:t', NS))
            start = 0
            while True:
                pos = joined.find(old, start)
                if pos < 0:
                    break
                paragraph_count += 1
                start = pos + 1
        if paragraph_count != count:
            raise ValueError("Target may span text nodes or occur outside a paragraph; no output written")
        if count == 0:
            split = any(old in "".join((t.text or "") for t in p.findall(".//w:t", NS))
                        for p in root.findall(".//w:p", NS))
            hint = "; the text appears to span multiple runs, which this safe filler does not edit" if split else ""
            raise ValueError(f"Text not found exactly within a text node: {old!r}{hint}")
        if count != 1 and not replace_all:
            raise ValueError(f"Text occurs {count} times, not once: {old!r}; specify --all only if every occurrence is intended")
        report.append({"old": old, "new": new, "occurrences": count})

    for spans in spans_by_node:
        spans.sort()
        if any(right[0] < left[1] for left, right in zip(spans, spans[1:])):
            raise ValueError("Replacement targets overlap in the source; no output written")

    # Apply every replacement to the original text, never to another new value.
    replacements = dict(pairs)
    pattern = re.compile("|".join(re.escape(old) for old, _ in pairs))

    def sub(m: re.Match[bytes]) -> bytes:
        decoded = html.unescape(m.group(2).decode("utf-8"))
        updated = pattern.sub(lambda hit: replacements[hit.group(0)], decoded)
        if updated == decoded:
            return m.group(0)
        opening = m.group(1)
        if updated[:1].isspace() or updated[-1:].isspace():
            if b"xml:space=" not in opening:
                opening = opening[:-1] + b" " + XML_SPACE + b">"
        return opening + html.escape(updated, quote=False).encode("utf-8") + m.group(3)

    return TEXT_RE.sub(sub, xml), report


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source", type=Path)
    ap.add_argument("--out", required=True, type=Path, help="new output DOCX; must not already exist")
    ap.add_argument("--replace", action="append", required=True, metavar="OLD=>NEW")
    ap.add_argument("--expect-sha256", help="expected source SHA-256 from occurrence inspection")
    ap.add_argument("--all", action="store_true", help="replace every exact occurrence (use only when explicitly intended)")
    args = ap.parse_args()
    source, out = args.source.resolve(), args.out.resolve()
    if not source.is_file():
        print(json.dumps({"error": "source_not_found", "path": str(source)}, ensure_ascii=False)); return 2
    if source == out:
        print(json.dumps({"error": "in_place_edit_refused"}, ensure_ascii=False)); return 2
    if out.exists():
        print(json.dumps({"error": "output_already_exists", "path": str(out)}, ensure_ascii=False)); return 2
    try:
        pairs = parse_pairs(args.replace)
        original_bytes = source.read_bytes()
        original_hash = hashlib.sha256(original_bytes).hexdigest()
        if args.expect_sha256 and args.expect_sha256.lower() != original_hash:
            raise ValueError("Source has changed since inspection; no output written")
        with zipfile.ZipFile(io.BytesIO(original_bytes), "r") as zin:
            bad = zin.testzip()
            if bad:
                raise ValueError(f"Corrupt ZIP member: {bad}")
            members = [(info, zin.read(info.filename)) for info in zin.infolist()]
            archive_comment = zin.comment
            index = next((i for i, (info, _) in enumerate(members) if info.filename == "word/document.xml"), None)
            if index is None:
                raise ValueError("word/document.xml not found")
        info, xml = members[index]
        changed_xml, report = apply_to_xml(xml, pairs, args.all)
        members[index] = (info, changed_xml)
        out.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp_name = tempfile.mkstemp(prefix=out.name + ".", suffix=".tmp", dir=out.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp_name, "w") as zout:
                zout.comment = archive_comment
                for member_info, data in members:
                    zout.writestr(member_info, data)
            with zipfile.ZipFile(tmp_name, "r") as check:
                bad = check.testzip()
                if bad:
                    raise ValueError(f"Output ZIP verification failed at {bad}")
                ET.fromstring(check.read("word/document.xml"))
            if out.exists():
                raise FileExistsError(f"Output appeared during operation; refusing to overwrite: {out}")
            if hashlib.sha256(source.read_bytes()).hexdigest() != original_hash:
                raise ValueError("Source changed during operation; no output written")
            os.rename(tmp_name, out)
        except Exception:
            if os.path.exists(tmp_name):
                os.unlink(tmp_name)
            raise
        print(json.dumps({
            "source": str(source), "source_sha256": original_hash,
            "output": str(out), "output_sha256": hashlib.sha256(out.read_bytes()).hexdigest(),
            "replacement_operations": report, "source_modified": hashlib.sha256(source.read_bytes()).hexdigest() != original_hash,
            "scope": "Only text in word/document.xml w:t nodes was changed; all other package parts preserved.",
        }, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(json.dumps({"error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False)); return 2


if __name__ == "__main__":
    sys.exit(main())
