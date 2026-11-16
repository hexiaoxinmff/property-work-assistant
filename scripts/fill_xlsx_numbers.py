#!/usr/bin/env python3
"""Patch existing numeric XLSX cells with explicit old/new values and preserve all other ZIP bytes."""
from pathlib import Path,PurePosixPath
from decimal import Decimal,InvalidOperation
from xml.etree import ElementTree as ET
import argparse,hashlib,json,os,posixpath,re,sys,tempfile,zipfile

M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P='{http://schemas.openxmlformats.org/package/2006/relationships}'
CELL=re.compile(rb'<c\b[^>]*\br="([A-Z]+[1-9][0-9]*)"[^>]*>.*?</c>',re.DOTALL)
VALUE=re.compile(rb'(<v>)([^<]+)(</v>)')
def num(v):
    if not isinstance(v,str):raise ValueError('期望值和新值必须为数字字符串')
    try:d=Decimal(v)
    except InvalidOperation:raise ValueError('非数字值')
    if not d.is_finite() or abs(d)>Decimal('1000000000000') or d.as_tuple().exponent < -8:raise ValueError('数字超出支持范围')
    return d
def main():
    a=argparse.ArgumentParser(description=__doc__);a.add_argument('source');a.add_argument('--sheet',required=True);a.add_argument('--plan',required=True);a.add_argument('--out',required=True);a=a.parse_args()
    source=Path(a.source).resolve();out=Path(a.out).resolve()
    if source==out or out.exists():raise ValueError('必须另存到不存在的文件')
    if source.suffix.lower()!='.xlsx' or out.suffix.lower()!='.xlsx':raise ValueError('仅支持无宏 XLSX')
    plan=json.loads(Path(a.plan).read_text(encoding='utf-8-sig'))
    if not isinstance(plan,list) or not plan:raise ValueError('修改计划必须是非空列表')
    desired={}
    for item in plan:
        addr=item.get('cell')
        if not isinstance(addr,str) or not re.fullmatch(r'[A-Z]+[1-9][0-9]*',addr):raise ValueError('单元格地址无效')
        if addr in desired:raise ValueError('单元格重复指定')
        desired[addr]=(num(item.get('expected')),num(item.get('value')))
    raw=source.read_bytes();original_hash=hashlib.sha256(raw).hexdigest()
    with zipfile.ZipFile(source) as z:
        if z.testzip():raise ValueError('源文件损坏')
        names=z.namelist()
        if len(names)!=len(set(names)):raise ValueError('存在重名条目')
        if 'xl/vbaProject.bin' in names:raise ValueError('不支持宏')
        if any(n.startswith('_xmlsignatures/') for n in names):raise ValueError('不支持已数字签名的工作簿')
        entries=[(i,z.read(i.filename)) for i in z.infolist()]
    data={i.filename:b for i,b in entries};wb=ET.fromstring(data['xl/workbook.xml']);relations=ET.fromstring(data['xl/_rels/workbook.xml.rels'])
    protection=wb.find(M+'workbookProtection')
    if protection is not None and any(str(v).lower() not in ('0','false','') for v in protection.attrib.values()):raise ValueError('不处理受保护的工作簿')
    sheets=[x for x in wb.findall('.//'+M+'sheet') if x.get('name')==a.sheet]
    if len(sheets)!=1:raise ValueError('工作表不存在或名称不唯一')
    rid=sheets[0].get(R+'id');links=[x for x in relations.findall(P+'Relationship') if x.get('Id')==rid]
    if len(links)!=1 or links[0].get('TargetMode')=='External':raise ValueError('工作表关联无效')
    loc=links[0].get('Target','')
    sheetpath=posixpath.normpath(loc.lstrip('/') if loc.startswith('/') else 'xl/'+loc)
    if not sheetpath.startswith('xl/worksheets/') or sheetpath not in data:raise ValueError('工作表文件不受支持')
    for path,content in data.items():
        if path.startswith('xl/worksheets/') and path.endswith('.xml'):
            node=ET.fromstring(content)
            if any(True for _ in node.iter(M+'f')):raise ValueError('工作簿含公式，须使用能重算的表格工具并另行验证')
    xml=data[sheetpath];tree=ET.fromstring(xml)
    if any(True for _ in tree.iter(M+'sheetProtection')):raise ValueError('不处理受保护的工作表')
    if any(True for _ in tree.iter(M+'mergeCell')):raise ValueError('工作表含合并单元格，当前工具不处理')
    if any(True for _ in tree.iter(M+'dataValidation')):raise ValueError('工作表含数据验证，当前工具不处理')
    changes={};matches=list(CELL.finditer(xml))
    for address,(expected,new) in desired.items():
        found=[m for m in matches if m.group(1).decode()==address]
        if len(found)!=1:raise ValueError('单元格缺失、重复或 XML 形式不支持：'+address)
        match=found[0];cell=match.group(0);root=ET.fromstring(cell)
        if root.get('t') not in (None,'n') or root.find(M+'f') is not None:raise ValueError('目标不是非公式数值单元格：'+address)
        values=list(VALUE.finditer(cell))
        if len(values)!=1:raise ValueError('目标数值节点不支持：'+address)
        v=values[0]
        try:actual=Decimal(v.group(2).decode('ascii'))
        except (InvalidOperation,UnicodeError):raise ValueError('原值不是有效数字：'+address)
        if not actual.is_finite() or actual!=expected:raise ValueError('原值与预期不符：'+address)
        formatted=format(new,'f').encode('ascii')
        updated=cell[:v.start(2)]+formatted+cell[v.end(2):]
        changes[(match.start(),match.end())]=updated
    newxml=xml
    for (start,end),updated in sorted(changes.items(),reverse=True):newxml=newxml[:start]+updated+newxml[end:]
    ET.fromstring(newxml)
    out.parent.mkdir(parents=True,exist_ok=True);fd,tmp=tempfile.mkstemp(dir=out.parent,prefix='xlsx-',suffix='.tmp');os.close(fd)
    try:
        with zipfile.ZipFile(tmp,'w') as z:
            for info,content in entries:z.writestr(info,newxml if info.filename==sheetpath else content)
        with zipfile.ZipFile(tmp) as z:
            if z.testzip():raise ValueError('输出损坏')
            if set(z.namelist())!=set(data):raise ValueError('文件成员变化')
            for name in names:
                if z.read(name)!=(newxml if name==sheetpath else data[name]):raise ValueError('存在计划外变化')
        if hashlib.sha256(source.read_bytes()).hexdigest()!=original_hash:raise ValueError('源文件在处理中发生变化')
        if out.exists():raise FileExistsError('输出已存在')
        os.rename(tmp,out)
    finally:
        if os.path.exists(tmp):os.unlink(tmp)
    print(json.dumps({'output':str(out),'source_sha256':original_hash,'sheet':a.sheet,'cells':list(desired),'other_package_parts_identical':True,'note':'仅核对 XML 和指定数值；图表、渲染、数据用途仍须复核。不处理含公式、合并单元格或数据验证的工作表。'},ensure_ascii=False))
if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
