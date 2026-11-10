#!/usr/bin/env python3
"""Locate text within DOCX body paragraphs; read-only, no changes to files."""
import argparse
import hashlib
import io
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('source', type=Path)
    p.add_argument('--text', required=True, help='exact text to locate')
    args = p.parse_args()
    if not args.text:
        print('Empty text is not permitted', file=sys.stderr)
        return 2
    try:
        raw = args.source.read_bytes()
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            xml = ET.fromstring(z.read('word/document.xml'))
        body = xml.find(W + 'body')
        if body is None:
            raise ValueError('Document body not found')
        hits = []
        for i, paragraph in enumerate(body.iter(W + 'p'), 1):
            nodes = list(paragraph.iter(W + 't'))
            text = ''.join(n.text or '' for n in nodes)
            ranges = []
            offset = 0
            for node in nodes:
                end = offset + len(node.text or '')
                ranges.append((offset, end))
                offset = end
            start = 0
            while True:
                pos = text.find(args.text, start)
                if pos < 0:
                    break
                hits.append({'paragraph': i, 'character_offset_zero_based': pos,
                    'context': text[max(0, pos-35):min(len(text), pos+len(args.text)+35)],
                    'contained_in_one_text_node': any(lo <= pos and pos + len(args.text) <= hi for lo, hi in ranges)})
                start = pos + 1
        print(json.dumps({'source': str(args.source), 'source_sha256': hashlib.sha256(raw).hexdigest(),
                          'occurrences': hits, 'count': len(hits),
                          'limitations': ['Body paragraphs and table paragraphs only; other document parts, page positions and semantic equivalence are not fully covered.',
                              'A node may contain multiple matches; location hints do not authorize replacement.',
                              'Output includes sensitive document context and may enter platform logs.']}, ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError, ET.ParseError, zipfile.BadZipFile, KeyError) as exc:
        print(f'Cannot inspect: {type(exc).__name__}: {exc}', file=sys.stderr)
        return 2

if __name__ == '__main__':
    raise SystemExit(main())
