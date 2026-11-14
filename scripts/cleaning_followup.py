#!/usr/bin/env python3
"""Summarize explicitly recorded cleaning events without claiming unrecorded inspections."""
from pathlib import Path
from datetime import datetime
from collections import Counter, defaultdict
import argparse,hashlib,json,sys
from openpyxl import load_workbook, Workbook
from xlsx_annotations import ensure_comments_safe
from xlsx_values import value_rows_without_formulas

FIELDS=['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明']
STATUS={'发现':'待处理','已通知':'已通知待反馈','反馈已完成':'待复查','复查通过':'已复查通过','复查不通过':'待整改'}
NEXT_ACTION={'待处理':'跟进处理安排','已通知待反馈':'跟进反馈，确认整改安排','待复查':'安排复查并记录结果','已复查通过':'留存记录','待整改':'跟进重新整改及复查','待核对':'先核实记录与状态'}
def text(v):return v.strip() if isinstance(v,str) else v
def ident(v):
    v=text(v);return v if isinstance(v,str) and v else None
def timestamp(v):
    if isinstance(v,datetime):return v if v.tzinfo is None else None
    if not isinstance(v,str):return None
    try:
        d=datetime.fromisoformat(v.strip())
        return d if d.tzinfo is None else None
    except ValueError:return None
def day_precision(v):
    if isinstance(v,datetime):
        return v.tzinfo is None and v.hour==0 and v.minute==0 and v.second==0 and v.microsecond==0
    if not isinstance(v,str):return False
    s=v.strip()
    return len(s)==10 and s[4:5]=='-' and s[7:8]=='-' and timestamp(s) is not None
def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('input');p.add_argument('--out',required=True);a=p.parse_args()
    inp=Path(a.input).resolve();out=Path(a.out).resolve()
    if inp==out or out.exists():raise ValueError('必须另存到不存在的输出路径')
    if inp.suffix.lower()!='.xlsx' or out.suffix.lower()!='.xlsx':raise ValueError('仅支持无宏 XLSX')
    original_hash=hashlib.sha256(inp.read_bytes()).hexdigest()
    ensure_comments_safe(inp)
    w=load_workbook(inp,read_only=True,data_only=False)
    try:
        if '巡检记录' not in w.sheetnames:raise ValueError('缺少巡检记录工作表')
        for sheet in w:
            if sheet.title!='巡检记录' and any(v is not None and (not isinstance(v,str) or v.strip()) for row in sheet.iter_rows(values_only=True) for v in row):
                raise ValueError('存在未识别且有内容的工作表，须先核实其业务影响：'+sheet.title)
        it=value_rows_without_formulas(w['巡检记录']);head=next(it,None)
        if head is None:raise ValueError('缺少表头')
        head=[text(v) for v in head]
        for field in FIELDS:
            if head.count(field)!=1:raise ValueError('表头缺失或重复：'+field)
        idx=[head.index(x) for x in FIELDS];records=[]
        extras=[i for i in range(len(head)) if i not in idx]
        named=[str(head[i]) for i in extras if head[i] is not None and head[i]!='']
        if named:raise ValueError('巡检记录存在未识别字段，须先核实映射和口径: '+', '.join(named))
        for rn,r in enumerate(it,2):
            if any(i<len(r) and r[i] is not None and r[i]!='' for i in extras):
                raise ValueError(f'巡检记录第{rn}行存在无表头数据，须先核实')
            if all(v is None for v in r):continue
            records.append((rn,{k:(r[i] if i<len(r) else None) for k,i in zip(FIELDS,idx)}))
    finally:w.close()
    ids=Counter(ident(r['记录编号']) for _,r in records if ident(r['记录编号']))
    grouped=defaultdict(list);issues=[];global_rows=[]
    for rn,r in records:
        rid=ident(r['记录编号']);iid=ident(r['事项编号']);errors=[]
        if not rid:errors.append('记录编号缺失或非文本')
        elif ids[rid]>1:errors.append('记录编号重复')
        if not iid:errors.append('事项编号缺失或非文本');global_rows.append(rn)
        if timestamp(r['记录时间']) is None:errors.append('记录时间无效或缺失')
        if text(r['事件']) not in STATUS:errors.append('事件不在已确认的测试事件类型内')
        if iid:grouped[iid].append((rn,r,errors))
        if errors:issues.append([rn,rid or '',iid or '', '；'.join(errors)])
    result=[]
    for iid,items in grouped.items():
        errors=[]
        if global_rows:errors.append('存在未归属记录，影响范围待确认')
        for rn,r,e in items:
            if e:errors.append('记录第'+str(rn)+'行待核实')
        dates=[timestamp(r['记录时间']) for _,r,_ in items]
        if len([x for x in dates if x is not None])!=len(set(x for x in dates if x is not None)):errors.append('同一事项存在相同记录时间，顺序待确认')
        for day in {d.date() for d in dates if d is not None}:
            same_day=[r for _,r,_ in items if timestamp(r['记录时间']) is not None and timestamp(r['记录时间']).date()==day]
            if len(same_day)>1 and any(day_precision(r['记录时间']) for r in same_day):
                errors.append('同一日期存在仅精确到日期的记录，事件先后待确认')
                break
        locations=[text(r['位置']) for _,r,_ in items if isinstance(text(r['位置']),str) and text(r['位置'])]
        if len(set(locations))>1:errors.append('同一事项位置记录不一致，须核实是否为同一事项')
        ordered=sorted(items,key=lambda t:timestamp(t[1]['记录时间']) or datetime.min)
        if not any(text(r['事件'])=='发现' for _,r,_ in ordered):errors.append('缺少发现记录')
        def latest(field):
            for _,r,_ in reversed(ordered):
                v=text(r[field])
                if v is not None and v!='':return v
            return '待确认'
        for field in ['位置','问题']:
            if latest(field)=='待确认':errors.append(field+'待确认')
        state='待核对' if errors else STATUS[text(ordered[-1][1]['事件'])]
        def recorded(field):
            if errors:return '待确认','待确认'
            for rn,r,_ in reversed(ordered):
                value=text(r[field])
                if value is not None and value!='':return value,'巡检记录!'+str(rn)
            return '待确认','未提供'
        owner,owner_source=recorded('负责人')
        deadline,deadline_source=recorded('期限')
        latest_event='待确认' if errors else text(ordered[-1][1]['事件'])
        latest_note='待确认' if errors else text(ordered[-1][1]['说明'])
        if latest_note is None or latest_note=='':latest_note='未提供'
        def stable(field):
            if not errors:return latest(field)
            values=[text(r[field]) for _,r,_ in items if text(r[field]) is not None and text(r[field])!='']
            return values[0] if values and all(v==values[0] for v in values) else '待确认'
        result.append([iid,stable('位置'),stable('问题'),owner,str(deadline),state,
                       latest_event,str(latest_note), '；'.join(errors),
                       ', '.join('巡检记录!'+str(rn) for rn,_,_ in ordered),NEXT_ACTION[state],
                       owner_source,deadline_source])
    if hashlib.sha256(inp.read_bytes()).hexdigest()!=original_hash:raise ValueError('源文件在处理时发生变化，未输出')
    out.parent.mkdir(parents=True,exist_ok=True);b=Workbook();ws=b.active;ws.title='整改跟进'
    ws.append(['事项编号','位置','问题','最近记录负责人（需核实当前）','最近记录期限（需核实当前）','当前状态','最新事件','最新事件说明','待核对原因','来源行','建议下一步（待人工确认）','负责人来源行','期限来源行'])
    for row in result:
        ws.append(row)
        for cell in ws[ws.max_row]:
            if isinstance(cell.value,str):cell.data_type='s'
    sheet=b.create_sheet('异常记录');sheet.append(['来源行','记录编号','事项编号','原因'])
    for row in issues:
        sheet.append(row)
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value,str):cell.data_type='s'
    note=b.create_sheet('说明');note.append(['注意','事件及确认状态仅依据输入记录，现场情况须按实际核对；待核对事项不得视为完成。'])
    for sheet in b:
        sheet.freeze_panes='A2'
        for col in 'ABCDEFGHIJKLM':sheet.column_dimensions[col].width=22
    with out.open('xb') as f:b.save(f)
    print(json.dumps({'output':str(out),'source_sha256':original_hash,'event_rows':len(records),'items':len(result),'review':sum(r[5]=='待核对' for r in result),'unassigned':len(global_rows)},ensure_ascii=False))
if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
