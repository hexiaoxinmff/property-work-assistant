#!/usr/bin/env python3
"""Simple local office files and persistent preference storage for the property assistant."""
import argparse
import datetime as dt
import json
import os
from pathlib import Path
import tempfile
import uuid


def as_json(value):
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    return str(value)


def emit(value):
    print(json.dumps(value, ensure_ascii=False, default=as_json, indent=2))


def check_new(path):
    path = Path(path)
    if path.exists():
        raise FileExistsError(f"输出文件已存在，拒绝覆盖: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def read_spec(path):
    with open(path, 'r', encoding='utf-8-sig') as f:
        spec = json.load(f)
    if not isinstance(spec, dict):
        raise ValueError('输入 JSON 须为对象')
    return spec


def read_docx(args):
    from docx import Document
    d = Document(args.input)
    emit({'paragraphs': [p.text for p in d.paragraphs],
          'tables': [[[c.text for c in row.cells] for row in t.rows] for t in d.tables],
          'limitations': '段落与表格分别输出，未完整读取页眉页脚、批注、文本框、图片文字和页面布局；合并单元格可能重复。'})


def create_docx(args):
    from docx import Document
    spec = read_spec(args.spec)
    dest = check_new(args.output)
    doc = Document()
    if spec.get('title'):
        doc.add_heading(str(spec['title']), 0)
    for text in spec.get('paragraphs', []):
        doc.add_paragraph(str(text))
    for rows in spec.get('tables', []):
        if not rows: continue
        width = max(len(row) for row in rows)
        if not width: continue
        table = doc.add_table(rows=0, cols=width)
        for row in rows:
            cells = table.add_row().cells
            for j, value in enumerate(row): cells[j].text = str(value if value is not None else '')
    with dest.open('xb') as f: doc.save(f)
    emit({'output': str(dest), 'note': '新建基础文档；不代表复刻原模板或已验证页面排版。'})


def read_xlsx(args):
    from openpyxl import load_workbook
    if min(args.start_row, args.start_col, args.max_rows, args.max_cols) < 1:
        raise ValueError('起始行列和读取数量必须大于零')
    workbook = load_workbook(args.input, read_only=True, data_only=False)
    try:
        output=[]
        for ws in workbook:
            if args.sheet and ws.title != args.sheet: continue
            rows=[]
            end_row=min(ws.max_row, args.start_row + args.max_rows - 1)
            end_col=min(ws.max_column, args.start_col + args.max_cols - 1)
            if args.start_row <= end_row and args.start_col <= end_col:
                for row in ws.iter_rows(min_row=args.start_row, max_row=end_row, min_col=args.start_col, max_col=end_col, values_only=True):
                    rows.append(list(row))
            output.append({'sheet':ws.title,'total_rows':ws.max_row,'total_columns':ws.max_column,
                           'start_row':args.start_row,'end_row':end_row if rows else None,
                           'start_col':args.start_col,'end_col':end_col if rows else None,
                           'more_rows':end_row < ws.max_row,'more_columns':end_col < ws.max_column,'rows':rows})
        if args.sheet and not output: raise ValueError('未找到指定工作表')
        emit({'sheets':output,'limitations':'保留公式文本，未计算公式；仅读取指定范围，图表、图片、样式和部分复杂结构未列出。'})
    finally: workbook.close()


def create_xlsx(args):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    spec=read_spec(args.spec); dest=check_new(args.output)
    sheets=spec.get('sheets',[])
    if not sheets: raise ValueError('需要至少一个工作表')
    wb=Workbook();wb.remove(wb.active)
    for item in sheets:
        ws=wb.create_sheet(str(item['name']))
        for row in item.get('rows',[]):
            ws.append([None]*len(row))
            for col,value in enumerate(row,1):
                if isinstance(value,(dict,list)): raise ValueError('单元格须为简单数值、布尔值、文字或空值')
                cell=ws.cell(ws.max_row,col)
                cell.value=value
                if isinstance(value,str): cell.data_type='s'
        if ws.max_row:
            for c in ws[1]: c.font=Font(bold=True);c.fill=PatternFill('solid',fgColor='DCEAF7')
        ws.freeze_panes='A2'
        for idx in range(1,min(ws.max_column,40)+1):
            length=max((len(str(ws.cell(r,idx).value or '')) for r in range(1,min(ws.max_row,100)+1)),default=8)
            ws.column_dimensions[get_column_letter(idx)].width=min(45,max(12,length*1.3+2))
    with dest.open('xb') as f: wb.save(f)
    emit({'output':str(dest),'note':'新建基础表格；所有文字作为文字保存，不计算公式、不修改原表。'})


def inspect_image(args):
    from PIL import Image,ImageOps
    with Image.open(args.input) as image:
        image_format=image.format
        image=ImageOps.exif_transpose(image)
        emit({'format':image_format, 'width':image.width,'height':image.height,'mode':image.mode,
              'limitations':'仅获取图像属性；不读取或识别画面、照片文字及其含义。请使用支持图片理解或 OCR 的工具。'})


def create_image(args):
    from PIL import Image,ImageDraw,ImageFont
    spec=read_spec(args.spec);dest=check_new(args.output)
    if dest.suffix.lower()!='.png': raise ValueError('文字通知图片请使用 .png 输出')
    width=1000; height=1400
    image=Image.new('RGB',(width,height),'white');draw=ImageDraw.Draw(image)
    candidates=[Path(os.environ.get('WINDIR',r'C:\Windows'))/'Fonts'/'msyh.ttc',Path('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc')]
    fontpath=next((str(p) for p in candidates if p.exists()),None)
    if not fontpath: raise RuntimeError('未找到可用的中文字体，请在目标平台安装或指定可用字体后再试')
    titlefont=ImageFont.truetype(fontpath,48);bodyfont=ImageFont.truetype(fontpath,32)
    title=str(spec.get('title','通知'));body='\n'.join(str(x) for x in spec.get('paragraphs',[]))
    draw.rectangle((0,0,width,24),fill='#1b5875')
    lines=[]
    for para in (title+'\n'+body).split('\n'):
        font=titlefont if not lines else bodyfont
        current=''
        for ch in para:
            if draw.textlength(current+ch,font=font)>width-120 and current:
                lines.append(current);current=ch;font=bodyfont
            else: current+=ch
        lines.append(current)
    y=80
    for idx,line in enumerate(lines):
        font=titlefont if idx==0 else bodyfont; step=82 if idx==0 else 58
        if y+step>height-60:raise ValueError('内容超过图片容量，请缩短文字或拆分图片')
        draw.text((60,y),line,font=font,fill='#1b2b36');y+=step
    with dest.open('xb') as f:image.save(f,format='PNG')
    emit({'output':str(dest),'note':'按已提供文字排版的基础通知图片，非照片理解或生成式画图。'})


def store_path(args):
    raw=args.store or os.environ.get('PROPERTY_ASSISTANT_PREFERENCES')
    return Path(raw).expanduser() if raw else Path.home()/'.property-work-assistant'/'preferences.json'


def load_store(path):
    if not path.exists():return {'schema':1,'preferences':[]}
    with path.open('r',encoding='utf-8') as f:data=json.load(f)
    if data.get('schema')!=1 or not isinstance(data.get('preferences'),list):raise ValueError('偏好文件格式不受支持')
    return data


def save_store(path,data):
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=None
    try:
        with tempfile.NamedTemporaryFile('w',encoding='utf-8',dir=path.parent,delete=False) as f:
            tmp=Path(f.name);json.dump(data,f,ensure_ascii=False,indent=2);f.write('\n');f.flush();os.fsync(f.fileno())
        os.replace(tmp,path)
    finally:
        if tmp and tmp.exists():tmp.unlink()


def prefs(args):
    path=store_path(args);data=load_store(path)
    if args.action=='list':
        emit({'path':str(path),'exists':path.exists(),'preferences':data['preferences']});return
    if not args.confirmed: raise ValueError('保存或删除长期偏好须由使用者明确要求，并传入 --confirmed')
    if args.action=='add':
        if not args.text or not args.scope: raise ValueError('添加偏好需要 --text 和 --scope')
        entry={'id':uuid.uuid4().hex,'scope':args.scope,'preference':args.text,'date':dt.date.today().isoformat(),'status':'confirmed-by-user'}
        data['preferences'].append(entry);save_store(path,data);emit({'path':str(path),'saved':entry})
    elif args.action=='update':
        found=[p for p in data['preferences'] if p.get('id')==args.id]
        if len(found)!=1:raise ValueError('未找到唯一偏好，未写入文件')
        if not args.text:raise ValueError('更新偏好需要 --text')
        found[0]['preference']=args.text
        if args.scope:found[0]['scope']=args.scope
        found[0]['updated']=dt.date.today().isoformat()
        save_store(path,data);emit({'path':str(path),'updated':found[0]})
    else:
        old=data['preferences'];data['preferences']=[p for p in old if p.get('id')!=args.id]
        if len(data['preferences'])==len(old):raise ValueError('未找到该偏好，未写入文件')
        save_store(path,data);emit({'path':str(path),'deleted':args.id})


def parser():
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='cmd',required=True)
    for name,func in [('read-docx',read_docx),('read-xlsx',read_xlsx),('inspect-image',inspect_image)]:
        x=sub.add_parser(name);x.add_argument('input');x.set_defaults(func=func)
        if name=='read-xlsx':x.add_argument('--sheet');x.add_argument('--start-row',type=int,default=1);x.add_argument('--start-col',type=int,default=1);x.add_argument('--max-rows',type=int,default=100);x.add_argument('--max-cols',type=int,default=40)
    for name,func in [('create-docx',create_docx),('create-xlsx',create_xlsx),('create-image',create_image)]:
        x=sub.add_parser(name);x.add_argument('--spec',required=True);x.add_argument('--output',required=True);x.set_defaults(func=func)
    x=sub.add_parser('preferences');x.add_argument('action',choices=['list','add','update','delete']);x.add_argument('--store');x.add_argument('--text');x.add_argument('--scope');x.add_argument('--id');x.add_argument('--confirmed',action='store_true');x.set_defaults(func=prefs)
    return p


if __name__=='__main__':
    import sys
    try:
        a=parser().parse_args();a.func(a)
    except Exception as e:
        print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
