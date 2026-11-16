#!/usr/bin/env python3
"""Create a verified notice draft from an explicitly reviewed synthetic or authorized ledger."""
from pathlib import Path
from decimal import Decimal,InvalidOperation
from xml.etree import ElementTree as ET
import argparse,hashlib,json,os,shutil,subprocess,sys,tempfile,zipfile
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parent
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
FIELDS=['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','已确认收款','待缴余额','处理状态','待核对原因','账单来源','收款来源']
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def call(script,args):
    env=os.environ.copy();env['PYTHONIOENCODING']='utf-8'
    r=subprocess.run([sys.executable,str(ROOT/script),*map(str,args)],capture_output=True,text=True,encoding='utf-8',env=env,timeout=120)
    try:obj=json.loads(r.stdout.strip() or r.stderr.strip())
    except ValueError:raise ValueError(script+' 未返回可核对结果')
    if r.returncode!=0:raise ValueError(script+' 未通过：'+str(obj.get('message',obj.get('error','未知原因'))))
    return obj
def reject_external_placeholders(path,targets):
    with zipfile.ZipFile(path) as z:
        names=z.namelist()
        if len(names)!=len(set(names)):raise ValueError('模板存在重复部件')
        if any(name.startswith('_xmlsignatures/') or name.endswith('vbaProject.bin') for name in names):raise ValueError('不支持含数字签名或宏的模板')
        for name in names:
            if not name.startswith('word/') or not name.endswith('.xml') or name=='word/document.xml':continue
            root=ET.fromstring(z.read(name))
            for p in root.iter(W+'p'):
                joined=''.join(t.text or '' for t in p.iter(W+'t'))
                if any(x in joined for x in targets):raise ValueError('页眉、页脚或其他正文外部件含目标占位符，无法确保全部填写')
def main():
    a=argparse.ArgumentParser(description=__doc__)
    a.add_argument('--ledger',required=True);a.add_argument('--bill-id',required=True);a.add_argument('--template',required=True);a.add_argument('--out',required=True)
    args=a.parse_args();ledger=Path(args.ledger).resolve();template=Path(args.template).resolve();out=Path(args.out).resolve()
    bill=args.bill_id.strip()
    if not bill:raise ValueError('账单编号不能为空')
    if out in (ledger,template) or out.exists():raise ValueError('必须另存到不存在的新文件')
    if ledger.suffix.lower()!='.xlsx' or template.suffix.lower()!='.docx' or out.suffix.lower()!='.docx':raise ValueError('需要 XLSX 台账和 DOCX 模板及输出')
    ledger_hash=sha(ledger);template_hash=sha(template)
    keys=['【商户名称】','【费用期间】','【费用类型】','【金额】']
    reject_external_placeholders(template,keys)
    out.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='property-notice-',dir=out.parent) as td:
        td=Path(td);review=td/'review.xlsx';draft=td/'draft.docx';plan=td/'plan.json'
        audit=call('arrears_review.py',[ledger,'--out',review])
        if audit.get('source_sha256')!=ledger_hash:raise ValueError('账单来源与预期不一致')
        w=load_workbook(review,read_only=True,data_only=False)
        try:
            it=w['跟进清单'].iter_rows(values_only=True);head=next(it,None)
            if list(head or [])!=FIELDS:raise ValueError('核对结果格式不受支持')
            matches=[r for r in it if r[0]==bill]
        finally:w.close()
        if len(matches)!=1:raise ValueError('账单不存在或编号不唯一')
        row=matches[0]
        if row[8]!='需跟进' or row[7] is None:raise ValueError('账单非已确认待跟进状态，不生成函件')
        try:balance=Decimal(str(row[7]))
        except InvalidOperation:raise ValueError('余额无效')
        if not balance.is_finite() or balance<=0 or balance!=balance.quantize(Decimal('0.01')):raise ValueError('余额无效')
        values=[row[2],row[3],row[4],format(balance,'.2f')]
        if any(not isinstance(v,str) or not v.strip() for v in values):raise ValueError('函件关键内容缺失')
        replacements=[{'old':k,'new':v} for k,v in zip(keys,values)]
        plan.write_text(json.dumps(replacements,ensure_ascii=False),encoding='utf-8')
        fillargs=[template,'--out',draft,'--expect-sha256',template_hash]
        for item in replacements:fillargs+=['--replace',item['old']+'=>'+item['new']]
        call('fill_docx.py',fillargs)
        checked=call('verify_docx_expected.py',[template,draft,'--plan',plan])
        if checked.get('passed') is not True:raise ValueError('修改计划核对未通过')
        if sha(ledger)!=ledger_hash or sha(template)!=template_hash:raise ValueError('源资料处理期间发生变化，未输出')
        with draft.open('rb') as source,out.open('xb') as dest:shutil.copyfileobj(source,dest)
    print(json.dumps({'output':str(out),'ledger_sha256':ledger_hash,'template_sha256':template_hash,'bill_id':bill,
                      'candidate_balance':str(balance),'verified_structure':True,'page_visual_verified':False,
                      'note':'草稿未发送；付款期限、收款方式和联系人如未提供应保留待确认。财务口径及页面须人工核对。'},ensure_ascii=False))
if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
