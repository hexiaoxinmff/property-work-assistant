#!/usr/bin/env python3
"""Reject workbook comments unless they are unaltered guidance in generated headers."""
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET
from zipfile import ZipFile
import posixpath,re
from make_blank_forms import COMMENTS

REL='{http://schemas.openxmlformats.org/package/2006/relationships}'
MAIN='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def ensure_comments_safe(path):
    with ZipFile(path) as z:
        names=z.namelist()
        if len(names)!=len(set(names)):raise ValueError('工作簿含重复部件，无法核验批注')
        known=set(names);candidates=set()
        for n in names:
            if n.lower().startswith('xl/') and n.lower().endswith('.xml'):
                basename=PurePosixPath(n).name.lower()
                if basename.startswith(('comments','comment','threadedcomment')):candidates.add(n)
            if not n.lower().startswith('xl/') or not n.lower().endswith('.rels'):continue
            root=ET.fromstring(z.read(n))
            if root.tag!=REL+'Relationships':raise ValueError('工作簿关系文件不受支持')
            for item in root.findall(REL+'Relationship'):
                typ=item.get('Type','').lower()
                if not ('comment' in typ and (typ.endswith('/comments') or typ.endswith('/threadedcomment'))):continue
                if item.get('TargetMode','').lower()=='external':raise ValueError('工作簿含外部批注，须核实')
                target=item.get('Target','')
                if target.startswith('/'):dest=posixpath.normpath(target.lstrip('/'))
                else:
                    parent=posixpath.dirname(posixpath.dirname(n))
                    dest=posixpath.normpath(posixpath.join(parent,target))
                if not dest.startswith('xl/') or dest not in known:raise ValueError('批注关联无法核实')
                candidates.add(dest)
        for name in sorted(candidates):
            if 'threadedcomment' in name.lower():raise ValueError('工作簿含讨论批注，须先人工核实')
            root=ET.fromstring(z.read(name))
            if root.tag!=MAIN+'comments':raise ValueError('工作簿批注结构不受支持')
            for item in root.iter(MAIN+'comment'):
                ref=item.get('ref','')
                content=''.join(x.text or '' for x in item.iter(MAIN+'t'))
                if not re.fullmatch(r'[A-Z]+1',ref) or content not in COMMENTS.values():
                    raise ValueError('工作簿含非标准表头说明或数据批注，须先人工核实其业务影响')
