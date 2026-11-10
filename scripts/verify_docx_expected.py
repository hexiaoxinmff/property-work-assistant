#!/usr/bin/env python3
"""Compare DOCX package and body text against an explicit replacement plan."""
import argparse,copy,hashlib,json,re,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def read(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip(): raise ValueError('DOCX 压缩包损坏')
        names=z.namelist()
        if len(set(names))!=len(names):raise ValueError('压缩包存在重名条目')
        return {name:z.read(name) for name in names}
def normalized(root):
    root=copy.deepcopy(root)
    for node in root.iter(W+'t'): node.text='__TEXT__'
    return ET.tostring(root,encoding='utf-8')
def main():
    a=argparse.ArgumentParser(description=__doc__);a.add_argument('source');a.add_argument('output');a.add_argument('--plan',required=True);args=a.parse_args()
    source=Path(args.source);out=Path(args.output)
    plan=json.loads(Path(args.plan).read_text(encoding='utf-8-sig'))
    if not isinstance(plan,list) or not plan:raise ValueError('计划必须是非空列表')
    pairs=[]
    for entry in plan:
        old=entry['old'];new=entry['new']
        if not isinstance(old,str) or not old or not isinstance(new,str):raise ValueError('替换内容必须是非空原文及字符串新值')
        pairs.append((old,new))
    if len({x[0] for x in pairs})!=len(pairs):raise ValueError('替换目标重复')
    x,y=read(source),read(out)
    if 'word/document.xml' not in x or 'word/document.xml' not in y:raise ValueError('缺少正文 XML')
    aroot=ET.fromstring(x['word/document.xml']);broot=ET.fromstring(y['word/document.xml'])
    anodes=list(aroot.iter(W+'t'));bnodes=list(broot.iter(W+'t'))
    pattern=re.compile('|'.join(re.escape(old) for old,_ in pairs));mapping=dict(pairs)
    counts={old:sum((n.text or '').count(old) for n in anodes) for old,_ in pairs}
    for old,_ in pairs:
        para_count=sum(( ''.join(t.text or '' for t in p.iter(W+'t')) ).count(old) for p in aroot.iter(W+'p'))
        if counts[old]!=1 or para_count!=1:raise ValueError('原文字段不唯一、跨片段或不在正文段落: '+old)
    for node in anodes:
        spans=[]; text=node.text or ''
        for old,_ in pairs:
            pos=text.find(old)
            if pos>=0:spans.append((pos,pos+len(old)))
        spans.sort()
        if any(right[0]<left[1] for left,right in zip(spans,spans[1:])):raise ValueError('替换目标重叠')
    expected=[pattern.sub(lambda m:mapping[m.group(0)],n.text or '') for n in anodes]
    actual=[n.text or '' for n in bnodes]
    others=sorted(name for name in set(x)|set(y) if name!='word/document.xml' and x.get(name)!=y.get(name))
    result={'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'output_sha256':hashlib.sha256(out.read_bytes()).hexdigest(),
            'same_package_members':set(x)==set(y),'other_parts_identical':not others,'different_parts':others,
            'body_structure_and_attributes_identical':normalized(aroot)==normalized(broot),
            'body_text_matches_plan_exactly':expected==actual,'expected_replacements':counts,
            'limitation':'只核对包成员和正文 XML 与替换计划，不证明分页、换行和视觉位置；页眉页脚内容必须另行检查。'}
    result['passed']=all(result[k] for k in ['same_package_members','other_parts_identical','body_structure_and_attributes_identical','body_text_matches_plan_exactly'])
    print(json.dumps(result,ensure_ascii=False,indent=2));return 0 if result['passed'] else 1
if __name__=='__main__':
    try:sys.exit(main())
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(2)
