#!/usr/bin/env python3
"""Create an empty, explicitly scoped workbook for the optional sample tools."""
from pathlib import Path
import argparse,json,sys
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

SHEETS={
 'arrears':[
  ('账单',['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整']),
  ('收款',['收款编号','账单编号','收款金额','收款已确认'])],
 'cleaning':[
  ('巡检记录',['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明'])]
}
COMMENTS={
 '账单编号':'文本编号；只使用经确认的账单关联，不以相似名称猜测。',
 '商户编号':'文本编号；同名商户使用各自编号。',
 '费用期间':'明确费用所属期间；由使用者确认实际口径。',
 '应收金额':'非负数，最多两位小数；空白不是零。',
 '账单已确认':'仅账单经相关人员确认后填写“是”，不得根据假设填写。',
 '收款记录完整':'仅该账单收款覆盖范围经确认完整后填写“是”；没有记录不代表完整或没有收款。',
 '收款编号':'文本编号；重复编号须先核查，不能自动合并。',
 '收款金额':'非负数，最多两位小数；空白不是零。',
 '收款已确认':'仅收款经相关人员确认后填写“是”。',
 '记录编号':'每条记录使用唯一文本编号。',
 '事项编号':'同一事项的记录使用经确认的同一文本编号。',
 '记录时间':'提供明确时间。仅日期时，同日多条记录的先后不能自动确定。',
 '事件':'测试映射：发现、已通知、反馈已完成、复查通过、复查不通过；实际含义须由使用者确认。',
 '说明':'保留原始说明；字段之间若有矛盾仍需人工核实。'
}
def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--kind',required=True,choices=sorted(SHEETS));p.add_argument('--out',required=True)
    a=p.parse_args();out=Path(a.out).resolve()
    if out.suffix.lower()!='.xlsx':raise ValueError('输出须为 XLSX')
    if out.exists():raise FileExistsError('输出已存在，不覆盖')
    w=Workbook();w.remove(w.active)
    for name,headers in SHEETS[a.kind]:
        s=w.create_sheet(name);s.append(headers);s.freeze_panes='A2'
        s.sheet_properties.pageSetUpPr.fitToPage=True
        s.page_setup.fitToWidth=1;s.page_setup.fitToHeight=0
        s.page_setup.orientation='landscape' if len(headers)>6 else 'portrait'
        s.page_setup.paperSize=s.PAPERSIZE_A4
        for col,head in enumerate(headers,1):
            c=s.cell(1,col);c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='1F4E78');c.alignment=Alignment(vertical='center',wrap_text=True)
            if head in COMMENTS:c.comment=Comment(COMMENTS[head],'物业办公助手')
            s.column_dimensions[get_column_letter(col)].width=max(16,min(26,len(head)*2+8))
            if head in ('账单编号','商户编号','收款编号','费用期间','记录编号','事项编号','记录时间'):
                for row in range(2,202):s.cell(row,col).number_format='@'
        s.row_dimensions[1].height=28
    out.parent.mkdir(parents=True,exist_ok=True)
    with out.open('xb') as f:w.save(f)
    print(json.dumps({'output':str(out),'kind':a.kind,'note':'空白试用格式，不代表公司现行台账或确认口径。'},ensure_ascii=False))
if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
