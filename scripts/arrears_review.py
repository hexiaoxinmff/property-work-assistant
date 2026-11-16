#!/usr/bin/env python3
"""Review explicit invoice/payment associations in a standard workbook; never send notices."""
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import argparse, hashlib, json, sys
from openpyxl import load_workbook, Workbook
from xlsx_annotations import ensure_comments_safe
from xlsx_values import value_rows_without_formulas
from openpyxl.styles import Font, PatternFill

BILL_HEADERS=['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整']
PAY_HEADERS=['收款编号','账单编号','收款金额','收款已确认']

def clean(value):
    return value.strip() if isinstance(value,str) else value

def identifier(value):
    v=clean(value)
    return v if isinstance(v,str) and v else None

def amount(value):
    if value is None or isinstance(value,(bool,dict,list)) or (isinstance(value,str) and not value.strip()):return None
    try:d=Decimal(str(value).strip())
    except (InvalidOperation,ValueError):return None
    try:
        if not d.is_finite() or d<0 or d>Decimal('1000000000000') or d!=d.quantize(Decimal('0.01')):return None
    except InvalidOperation:return None
    return d

def rows(sheet,headers):
    it=value_rows_without_formulas(sheet);head=next(it,None)
    if head is None:raise ValueError('空表: '+sheet.title)
    names=[clean(x) for x in head]
    for h in headers:
        if names.count(h)!=1:raise ValueError(f'{sheet.title}缺少或存在重复列: {h}')
    indexes=[names.index(h) for h in headers]
    extras=[i for i in range(len(names)) if i not in indexes]
    named=[str(names[i]) for i in extras if names[i] is not None and names[i]!='']
    if named:raise ValueError(f'{sheet.title}存在未识别字段，须先核实映射和口径: '+', '.join(named))
    for number,row in enumerate(it,2):
        if any(i<len(row) and row[i] is not None and row[i]!='' for i in extras):
            raise ValueError(f'{sheet.title}第{number}行存在无表头数据，须先核实')
        if all(v is None for v in row):continue
        yield number,{h:(row[i] if i<len(row) else None) for h,i in zip(headers,indexes)}

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('input');ap.add_argument('--out',required=True);args=ap.parse_args()
    inp=Path(args.input).resolve();out=Path(args.out).resolve()
    if inp==out or out.exists():raise ValueError('必须另存到不存在的输出路径')
    if inp.suffix.lower()!='.xlsx' or out.suffix.lower()!='.xlsx':raise ValueError('仅支持无宏 XLSX')
    original_hash=hashlib.sha256(inp.read_bytes()).hexdigest()
    ensure_comments_safe(inp)
    book=load_workbook(inp,read_only=True,data_only=False)
    try:
        if not {'账单','收款'}<=set(book.sheetnames):raise ValueError('需要账单和收款工作表')
        for sheet in book:
            if sheet.title not in {'账单','收款'} and any(v is not None and (not isinstance(v,str) or v.strip()) for row in sheet.iter_rows(values_only=True) for v in row):
                raise ValueError('存在未识别且有内容的工作表，须先核实其业务影响：'+sheet.title)
        bills=list(rows(book['账单'],BILL_HEADERS));payments=list(rows(book['收款'],PAY_HEADERS))
    finally:book.close()
    bill_counts=Counter(identifier(x['账单编号']) for _,x in bills if identifier(x['账单编号']))
    pay_counts=Counter(identifier(x['收款编号']) for _,x in payments if identifier(x['收款编号']))
    bybill=defaultdict(list);issues=[];unallocated=[]
    for row,p in payments:
        bid=identifier(p['账单编号']);pid=identifier(p['收款编号'])
        reasons=[]
        if not pid:reasons.append('收款编号缺失或非文本')
        elif pay_counts[pid]>1:reasons.append('收款编号重复')
        if not bid:reasons.append('账单编号缺失或非文本')
        elif bill_counts[bid]!=1:reasons.append('账单编号未匹配或不唯一')
        if amount(p['收款金额']) is None:reasons.append('收款金额无效')
        if clean(p['收款已确认'])!='是':reasons.append('收款未确认')
        if bid and bill_counts[bid]==1:bybill[bid].append((row,p,reasons))
        else:unallocated.append(row)
        if reasons:issues.append(['收款',row,pid or '',bid or '', '；'.join(reasons)])
    output=[];known=Decimal('0.00')
    for row,b in bills:
        bid=identifier(b['账单编号']);tenant=identifier(b['商户编号']);name=clean(b['商户名称'])
        problems=[]
        if unallocated:problems.append('存在未能唯一关联的收款记录，其影响范围待确认')
        if not bid:problems.append('账单编号缺失或非文本')
        elif bill_counts[bid]>1:problems.append('账单编号重复')
        if not tenant:problems.append('商户编号缺失或非文本')
        if not isinstance(name,str) or not name:problems.append('商户名称缺失')
        for field in ['费用期间','费用类型']:
            if not isinstance(clean(b[field]),str) or not clean(b[field]):problems.append(field+'缺失或非文本')
        due=amount(b['应收金额'])
        if due is None:problems.append('应收金额无效')
        if clean(b['账单已确认'])!='是':problems.append('账单未确认')
        if clean(b['收款记录完整'])!='是':problems.append('收款记录完整性未确认')
        matched=bybill.get(bid,[]) if bid else []
        for r,p,errors in matched:
            if errors:problems.append('收款第'+str(r)+'行待核实')
        total=sum((amount(p['收款金额']) for _,p,e in matched if not e),Decimal('0.00'))
        if due is not None and not problems and total>due:problems.append('已确认收款超过应收，处理口径待确认')
        if problems:
            state='待核对';balance=None
            issues.append(['账单',row,bid or '',tenant or '', '；'.join(problems)])
        else:
            balance=due-total;state='已结清' if balance==0 else '需跟进';known+=balance
        output.append([bid or '',tenant or '',name or '',clean(b['费用期间']) or '',clean(b['费用类型']) or '',
                       float(due) if due is not None else None,float(total) if not problems else None,
                       float(balance) if balance is not None else None,state,'；'.join(problems),
                       f'账单!{row}', ', '.join(f'收款!{r}' for r,_,_ in matched) or '无关联记录'])
    result=Workbook();ws=result.active;ws.title='跟进清单'
    ws.append(['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','已确认收款','待缴余额','处理状态','待核对原因','账单来源','收款来源'])
    for r in output:
        ws.append(r)
        for c in ws[ws.max_row]:
            if isinstance(c.value,str):c.data_type='s'
    issue=result.create_sheet('异常记录');issue.append(['来源表','原行号','记录编号','关联编号','问题'])
    for r in issues:
        issue.append(r)
        for c in issue[issue.max_row]:
            if isinstance(c.value,str):c.data_type='s'
    eligible=sum(r[8]!='待核对' for r in output)
    summary=result.create_sheet('汇总');summary.append(['项目','结果']);summary.append(['账单行数',len(output)]);summary.append(['需跟进笔数',sum(r[8]=='需跟进' for r in output)]);summary.append(['待核对笔数',sum(r[8]=='待核对' for r in output)]);summary.append(['已确认待缴余额（不含待核对）',float(known) if eligible else None]);summary.append(['已纳入核对笔数',eligible]);summary.append(['异常记录数',len(issues)]);summary.append(['注意','待核对或未关联记录不包含在已确认余额中；无可核对记录时余额留空，本汇总不是全部欠费总额。'])
    for ws in result:
        ws.freeze_panes='A2'
        for cell in ws[1]:cell.font=Font(bold=True);cell.fill=PatternFill('solid',fgColor='DCEAF7')
        ws.column_dimensions['A'].width=23
        for col in 'BCDEFGHIJKL':ws.column_dimensions[col].width=21
    if hashlib.sha256(inp.read_bytes()).hexdigest()!=original_hash:raise ValueError('源文件在处理时发生变化，未输出')
    out.parent.mkdir(parents=True,exist_ok=True)
    with out.open('xb') as f:result.save(f)
    print(json.dumps({'output':str(out),'source_sha256':original_hash,'bill_rows':len(output),'review_rows':sum(r[8]=='待核对' for r in output),'issue_rows':len(issues),'confirmed_balance':str(known) if eligible else None,'note':'仅按明确关联和确认状态汇总；待核对项未包含在已确认余额中。'},ensure_ascii=False))

if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
