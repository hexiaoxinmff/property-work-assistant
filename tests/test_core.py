#!/usr/bin/env python3
"""Offline synthetic regression tests; no customer files, network, or external writes."""
import contextlib,hashlib,http.client,io,json,os,re,subprocess,sys,tempfile,unittest,zipfile
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parent.parent/'scripts'))
import discover_skills

ROOT=Path(__file__).resolve().parent.parent
SCRIPTS=ROOT/'scripts'
try:
    from openpyxl import Workbook,load_workbook
    EXCEL=True
except ImportError:EXCEL=False
try:
    from docx import Document
    WORD=True
except ImportError:WORD=False

class CoreTests(unittest.TestCase):
    def test_ocr_documentation_uses_portable_paths(self):
        documentation=(ROOT/'references'/'ocr-winrt.md').read_text(encoding='utf-8')
        self.assertIsNone(re.search(r'(?im)\b[A-Z]:\\Users\\[^\\]+\\',documentation))
        self.assertIn(r'".\scripts\ocr_images_winrt.ps1"',documentation)
        self.assertIn(r'".\scripts\ocr_pdf_winrt.ps1"',documentation)
    def setUp(self):self.temp=tempfile.TemporaryDirectory(prefix='property-test-');self.d=Path(self.temp.name)
    def tearDown(self):self.temp.cleanup()
    def run_tool(self,script,*args,success=True):
        env=os.environ.copy();env['PYTHONIOENCODING']='utf-8'
        r=subprocess.run([sys.executable,str(SCRIPTS/script),*map(str,args)],capture_output=True,text=True,encoding='utf-8',env=env)
        self.assertEqual(r.returncode==0,success,(r.stdout,r.stderr))
        return r
    def workbook(self,name,sheets):
        w=Workbook();w.remove(w.active)
        for title,rows in sheets.items():
            sh=w.create_sheet(title)
            for row in rows:sh.append(row)
        path=self.d/name;w.save(path);return path
    @unittest.skipUnless(EXCEL and WORD,'openpyxl or python-docx unavailable')
    def test_formula_business_values_never_become_confirmed_results(self):
        headers=['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整']
        payheaders=['收款编号','账单编号','收款金额','收款已确认']
        doc=Document();doc.add_paragraph('【商户名称】【费用期间】【费用类型】【金额】');template=self.d/'formula-template.docx';doc.save(template)
        for field,formula in [(0,'="B1"'),(2,'="虚构商户"'),(5,'=1000')]:
            row=['B1','T1','虚构商户','2026-10','物业费',1000,'是','是'];row[field]=formula
            src=self.workbook(f'formula-{field}.xlsx',{'账单':[headers,row],'收款':[payheaders]})
            out=self.d/f'formula-{field}-out.xlsx';r=self.run_tool('arrears_review.py',src,'--out',out,success=False)
            self.assertIn('公式',r.stderr);self.assertFalse(out.exists())
            notice=self.d/f'formula-{field}.docx';self.run_tool('arrears_notice.py','--ledger',src,'--bill-id','B1','--template',template,'--out',notice,success=False);self.assertFalse(notice.exists())
        fields=['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明']
        events=[fields,['R1','I1','2026-11-10 09:00','发现','="东门"','垃圾',None,None,''],['R2','I1','2026-11-10 10:00','复查通过',None,None,None,None,'']]
        src=self.workbook('formula-cleaning.xlsx',{'巡检记录':events});out=self.d/'formula-cleaning-out.xlsx'
        r=self.run_tool('cleaning_followup.py',src,'--out',out,success=False);self.assertIn('公式',r.stderr);self.assertFalse(out.exists())
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_arrears_confirmed_and_incomplete(self):
        bills=[['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整'],
               ['B1','T1','示例','2026-10','物业费',2400.05,'是','是'],
               ['B2','T2','示例','2026-09','物业费',500,'是','否']]
        pay=[['收款编号','账单编号','收款金额','收款已确认'],['P1','B1',1000.02,'是'],['P2','B1',500.03,'是']]
        src=self.workbook('input.xlsx',{'账单':bills,'收款':pay});old=hashlib.sha256(src.read_bytes()).digest();out=self.d/'result.xlsx'
        r=self.run_tool('arrears_review.py',src,'--out',out);data=json.loads(r.stdout);self.assertEqual(data['confirmed_balance'],'900.00')
        w=load_workbook(out,data_only=True);rows=list(w['跟进清单'].values);w.close()
        self.assertEqual((rows[1][7],rows[1][8]),(900,'需跟进'));self.assertEqual((rows[2][7],rows[2][8]),(None,'待核对'))
        self.assertEqual(hashlib.sha256(src.read_bytes()).digest(),old);self.run_tool('arrears_review.py',src,'--out',out,success=False)
    @unittest.skipUnless(EXCEL and WORD,'openpyxl or python-docx unavailable')
    def test_notice_pipeline_confirmed_and_held(self):
        bills=[['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整'],
               ['B1','T1','虚构商户','2026-10','物业费',2400.05,'是','是'],
               ['B2','T2','虚构商户乙','2026-10','物业费',200,'是','否'],
               ['B3','T3','虚构商户丙','2026-10','物业费',0,'是','是']]
        pays=[['收款编号','账单编号','收款金额','收款已确认'],['P1','B1',1000.02,'是'],['P2','B1',500.03,'是']]
        ledger=self.workbook('ledger.xlsx',{'账单':bills,'收款':pays})
        d=Document();d.add_paragraph('致：【商户名称】');d.add_paragraph('期间：【费用期间】')
        d.add_paragraph('费用：【费用类型】');d.add_paragraph('待缴：【金额】元')
        d.add_paragraph('期限：【付款期限】');template=self.d/'template.docx';d.save(template)
        before=(hashlib.sha256(ledger.read_bytes()).hexdigest(),hashlib.sha256(template.read_bytes()).hexdigest())
        out=self.d/'notice.docx';r=self.run_tool('arrears_notice.py','--ledger',ledger,'--bill-id','B1','--template',template,'--out',out)
        info=json.loads(r.stdout);self.assertTrue(info['verified_structure']);self.assertFalse(info['page_visual_verified'])
        content='\\n'.join(p.text for p in Document(out).paragraphs);self.assertIn('900.00',content);self.assertIn('【付款期限】',content)
        self.assertEqual(before,(hashlib.sha256(ledger.read_bytes()).hexdigest(),hashlib.sha256(template.read_bytes()).hexdigest()))
        self.run_tool('arrears_notice.py','--ledger',ledger,'--bill-id','B1','--template',template,'--out',out,success=False)
        for bid in ['B2','B3','missing']:
            blocked=self.d/(bid+'.docx');self.run_tool('arrears_notice.py','--ledger',ledger,'--bill-id',bid,'--template',template,'--out',blocked,success=False)
            self.assertFalse(blocked.exists())
        d.sections[0].header.paragraphs[0].text='【商户名称】'
        t2=self.d/'header-duplicate.docx';d.save(t2);blocked=self.d/'header-output.docx'
        self.run_tool('arrears_notice.py','--ledger',ledger,'--bill-id','B1','--template',t2,'--out',blocked,success=False)
        self.assertFalse(blocked.exists())
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_arrears_rejects_unknown_business_columns_and_unheaded_data(self):
        headers=['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整']
        pay=[['收款编号','账单编号','收款金额','收款已确认']]
        for extra_header in ['调减金额',None]:
            with self.subTest(extra_header=extra_header):
                bills=[headers+[extra_header],['B1','T1','示例','2026-10','物业费',1000,'是','是',200]]
                src=self.workbook(('named' if extra_header else 'unnamed')+'.xlsx',{'账单':bills,'收款':pay})
                out=self.d/(('named' if extra_header else 'unnamed')+'-out.xlsx')
                self.run_tool('arrears_review.py',src,'--out',out,success=False)
                self.assertFalse(out.exists())
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_cleaning_conflicting_locations_hold_completion(self):
        headers=['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明']
        rows=[headers,['R1','I1','2026-11-10 09:00','发现','东门','垃圾',None,None,None],['R2','I1','2026-11-10 10:00','复查通过','西门',None,None,None,None]]
        src=self.workbook('locations.xlsx',{'巡检记录':rows});out=self.d/'locations-out.xlsx'
        self.run_tool('cleaning_followup.py',src,'--out',out)
        w=load_workbook(out,data_only=True);item=list(w['整改跟进'].values)[1];w.close()
        self.assertEqual(item[1],'待确认');self.assertEqual(item[5],'待核对');self.assertIn('位置记录不一致',item[8]);self.assertEqual(item[10],'先核实记录与状态')
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_cleaning_feedback_requires_inspection(self):
        rows=[['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明'],
              ['R1','I1','2026-11-10 09:00','发现','车库','积水','保洁甲',None,''],
              ['R2','I1','2026-11-10 10:00','反馈已完成',None,None,None,None,'已反馈']]
        src=self.workbook('clean.xlsx',{'巡检记录':rows});out=self.d/'result.xlsx';self.run_tool('cleaning_followup.py',src,'--out',out)
        w=load_workbook(out,data_only=True);item=list(w['整改跟进'].values)[1];w.close();self.assertEqual(item[5],'待复查')
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_cleaning_latest_note_and_historical_owner_provenance(self):
        headers=['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明']
        rows=[headers,
              ['R1','I1','2026-11-10 09:00','发现','东门','垃圾','保洁甲','2026-11-10 12:00','现场发现垃圾'],
              ['R2','I1','2026-11-10 15:00','反馈已完成',None,None,None,None,None],
              ['R3','I2','2026-11-10','发现','北门','污渍','保洁乙',None,'旧记录'],
              ['R4','I2','2026-11-10 10:00','复查通过','南门',None,'保洁丙',None,'未核实顺序']]
        src=self.workbook('events.xlsx',{'巡检记录':rows});out=self.d/'events-out.xlsx'
        self.run_tool('cleaning_followup.py',src,'--out',out)
        w=load_workbook(out,data_only=True);values=list(w['整改跟进'].values);w.close()
        self.assertIn('需核实当前',values[0][3]);items={x[0]:x for x in values[1:]}
        self.assertEqual(items['I1'][6],'反馈已完成');self.assertEqual(items['I1'][7],'未提供')
        self.assertEqual(items['I1'][3],'保洁甲');self.assertEqual(items['I1'][11],'巡检记录!2')
        self.assertEqual(items['I1'][12],'巡检记录!2')
        self.assertEqual(items['I2'][5],'待核对');self.assertEqual(items['I2'][6],'待确认')
        self.assertEqual(items['I2'][7],'待确认');self.assertEqual(items['I2'][3],'待确认');self.assertEqual(items['I2'][1],'待确认')
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_cleaning_next_actions_follow_recorded_states(self):
        headers=['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明']
        events=['发现','已通知','反馈已完成','复查通过','复查不通过']
        expected=['跟进处理安排','跟进反馈，确认整改安排','安排复查并记录结果','留存记录','跟进重新整改及复查']
        rows=[headers]
        for n,event in enumerate(events,1):
            rows.append([f'R{n}a',f'I{n}','2026-11-10 09:00','发现','东门','虚构问题',None,None,''])
            if event!='发现':rows.append([f'R{n}b',f'I{n}','2026-11-10 10:00',event,None,None,None,None,''])
        rows.append(['R6a','I6','2026-11-10','发现','东门','虚构问题',None,None,''])
        rows.append(['R6b','I6','2026-11-10 10:00','复查通过',None,None,None,None,''])
        src=self.workbook('steps.xlsx',{'巡检记录':rows});out=self.d/'steps-out.xlsx'
        self.run_tool('cleaning_followup.py',src,'--out',out)
        w=load_workbook(out,data_only=True);records={r[0]:r for r in list(w['整改跟进'].values)[1:]};w.close()
        for n,action in enumerate(expected,1):self.assertEqual(records[f'I{n}'][10],action)
        self.assertEqual(records['I2'][5],'已通知待反馈')
        self.assertEqual(records['I6'][5],'待核对');self.assertEqual(records['I6'][10],'先核实记录与状态')
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_cleaning_rejects_unknown_and_unheaded_columns(self):
        headers=['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明']
        for extra in ['复查补充结果',None]:
            with self.subTest(extra=extra):
                rows=[headers+[extra],
                      ['R1','I1','2026-11-10 09:00','发现','东门','垃圾','保洁甲',None,'',''],
                      ['R2','I1','2026-11-10 10:00','复查通过',None,None,None,None,'','实际复查不通过']]
                label='named' if extra else 'unheaded'
                src=self.workbook(label+'.xlsx',{'巡检记录':rows});out=self.d/(label+'-output.xlsx')
                self.run_tool('cleaning_followup.py',src,'--out',out,success=False)
                self.assertFalse(out.exists())
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_arrears_extra_sheet_must_not_be_ignored(self):
        bills=[['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整'],
               ['B1','T1','示例','2026-10','物业费',1000,'是','是']]
        payments=[['收款编号','账单编号','收款金额','收款已确认']]
        src=self.workbook('extra.xlsx',{'账单':bills,'收款':payments,'调整记录':[['B1','调减200']]})
        out=self.d/'out.xlsx';self.run_tool('arrears_review.py',src,'--out',out,success=False);self.assertFalse(out.exists())
        blank=self.workbook('blank.xlsx',{'账单':bills,'收款':payments,'空表':[]})
        self.run_tool('arrears_review.py',blank,'--out',self.d/'blank-out.xlsx')
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_cleaning_extra_sheet_must_not_be_ignored(self):
        rows=[['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明'],
              ['R1','I1','2026-11-10 09:00','发现','车库','积水','保洁甲',None,''],
              ['R2','I1','2026-11-10 10:00','复查通过',None,None,None,None,'']]
        src=self.workbook('extra.xlsx',{'巡检记录':rows,'后续复查':[['I1','不通过']]})
        out=self.d/'out.xlsx';self.run_tool('cleaning_followup.py',src,'--out',out,success=False);self.assertFalse(out.exists())
        blank=self.workbook('blank.xlsx',{'巡检记录':rows,'空表':[]})
        self.run_tool('cleaning_followup.py',blank,'--out',self.d/'blank-out.xlsx')
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_cleaning_ambiguous_same_day_order(self):
        headers=['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明']
        for day_value in ['2026-11-10',__import__('datetime').datetime(2026,11,10)]:
            with self.subTest(day_value=str(day_value)):
                rows=[headers,['R1','I1',day_value,'发现','东门','垃圾','保洁甲',None,''],
                      ['R2','I1','2026-11-10 10:00','复查通过',None,None,None,None,'']]
                label='text' if isinstance(day_value,str) else 'exceldate'
                src=self.workbook(label+'.xlsx',{'巡检记录':rows});out=self.d/(label+'-out.xlsx')
                self.run_tool('cleaning_followup.py',src,'--out',out)
                w=load_workbook(out,data_only=True);item=list(w['整改跟进'].values)[1];w.close()
                self.assertEqual(item[5],'待核对');self.assertIn('先后待确认',item[8])
        rows=[headers,['R1','I1','2026-11-10','发现','东门','垃圾','保洁甲',None,''],
              ['R2','I1','2026-11-11 10:00','复查通过',None,None,None,None,'']]
        src=self.workbook('differentday.xlsx',{'巡检记录':rows});out=self.d/'differentday-out.xlsx'
        self.run_tool('cleaning_followup.py',src,'--out',out)
        w=load_workbook(out,data_only=True);item=list(w['整改跟进'].values)[1];w.close()
        self.assertEqual(item[5],'已复查通过')
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_blank_templates_do_not_invent_business_data(self):
        for kind,script in [('arrears','arrears_review.py'),('cleaning','cleaning_followup.py')]:
            with self.subTest(kind=kind):
                template=self.d/(kind+'.xlsx');out=self.d/(kind+'-out.xlsx')
                self.run_tool('make_blank_forms.py','--kind',kind,'--out',template)
                w=load_workbook(template);self.assertEqual(w.sheetnames,['账单','收款'] if kind=='arrears' else ['巡检记录'])
                for sheet in w:
                    self.assertTrue(all(v is None for row in sheet.iter_rows(min_row=2,values_only=True) for v in row))
                    self.assertEqual(sheet.page_setup.fitToWidth,1)
                    self.assertEqual(sheet.page_setup.fitToHeight,0)
                if kind=='arrears':
                    self.assertIsNotNone(w['账单']['H1'].comment)
                    self.assertEqual(w['账单']['A2'].number_format,'@')
                else:
                    self.assertIsNotNone(w['巡检记录']['D1'].comment)
                w.close()
                self.run_tool('make_blank_forms.py','--kind',kind,'--out',template,success=False)
                result=json.loads(self.run_tool(script,template,'--out',out).stdout)
                if kind=='arrears':
                    self.assertEqual(result['bill_rows'],0);self.assertIsNone(result['confirmed_balance'])
                else:self.assertEqual(result['event_rows'],0)
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_unknown_comments_block_automatic_results(self):
        from openpyxl.comments import Comment
        bills=[['账单编号','商户编号','商户名称','费用期间','费用类型','应收金额','账单已确认','收款记录完整'],
               ['B1','T1','示例','2026-10','物业费',1000,'是','是']]
        pay=[['收款编号','账单编号','收款金额','收款已确认']]
        cleaning=[['记录编号','事项编号','记录时间','事件','位置','问题','负责人','期限','说明'],
                  ['R1','I1','2026-11-10 09:00','发现','车库','积水','保洁甲',None,''],
                  ['R2','I1','2026-11-10 10:00','复查通过',None,None,None,None,'']]
        cases=[('data-bill',{'账单':bills,'收款':pay},'账单','F2','已审批调减200','arrears_review.py'),
               ('header-bill',{'账单':bills,'收款':pay},'账单','F1','金额包含调减','arrears_review.py'),
               ('data-clean',{'巡检记录':cleaning},'巡检记录','D3','复查实际不通过','cleaning_followup.py')]
        for name,sheets,sheet,cell,comment,tool in cases:
            with self.subTest(name=name):
                src=self.workbook(name+'.xlsx',sheets);w=load_workbook(src);w[sheet][cell].comment=Comment(comment,'虚构');w.save(src);w.close()
                out=self.d/(name+'-out.xlsx');self.run_tool(tool,src,'--out',out,success=False);self.assertFalse(out.exists())
    @unittest.skipUnless(EXCEL,'openpyxl unavailable')
    def test_excel_patch_and_formula_rejection(self):
        src=self.workbook('input.xlsx',{'台账':[['事项','金额'],['样例',1234.5]]});out=self.d/'result.xlsx'
        plan=self.d/'plan.json';plan.write_text(json.dumps([{'cell':'B2','expected':'1234.5','value':'900.25'}]),encoding='utf-8')
        self.run_tool('fill_xlsx_numbers.py',src,'--sheet','台账','--plan',plan,'--out',out)
        w=load_workbook(out);self.assertEqual(w['台账']['B2'].value,900.25);w.close()
        src2=self.workbook('formula.xlsx',{'台账':[['事项','金额','公式'],['样例',1234.5,'=B2*2']]});bad=self.d/'blocked.xlsx'
        self.run_tool('fill_xlsx_numbers.py',src2,'--sheet','台账','--plan',plan,'--out',bad,success=False);self.assertFalse(bad.exists())
    @unittest.skipUnless(WORD,'python-docx unavailable')
    def test_word_template_exact_and_split_rejection(self):
        d=Document();d.sections[0].header.paragraphs[0].text='虚构页眉'
        d.add_paragraph('商户：【商户】')
        t=d.add_table(rows=1,cols=1);t.cell(0,0).text='金额：【金额】'
        p=d.add_paragraph('跨片段：');p.add_run('【分');p.add_run('段】')
        src=self.d/'source.docx';d.save(src);before=hashlib.sha256(src.read_bytes()).hexdigest()
        out=self.d/'result.docx';plan=self.d/'plan.json'
        changes=[{'old':'【商户】','new':'虚构商户'},{'old':'【金额】','new':'900'}]
        plan.write_text(json.dumps(changes,ensure_ascii=False),encoding='utf-8')
        self.run_tool('fill_docx.py',src,'--out',out,'--expect-sha256',before,'--replace','【商户】=>虚构商户','--replace','【金额】=>900')
        r=self.run_tool('verify_docx_expected.py',src,out,'--plan',plan);self.assertTrue(json.loads(r.stdout)['passed'])
        blocked=self.d/'blocked.docx'
        self.run_tool('fill_docx.py',src,'--out',blocked,'--replace','【分段】=>新值',success=False)
        self.assertFalse(blocked.exists());self.assertEqual(hashlib.sha256(src.read_bytes()).hexdigest(),before)
    @unittest.skipUnless(WORD,'python-docx unavailable')
    def test_word_complex_template_preserves_nontext_parts(self):
        import struct,zlib
        from docx.shared import Inches
        from docx.enum.section import WD_SECTION
        def png_chunk(kind,data):
            return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)
        png=b'\x89PNG\r\n\x1a\n'+png_chunk(b'IHDR',struct.pack('>IIBBBBB',1,1,8,6,0,0,0))+png_chunk(b'IDAT',zlib.compress(b'\x00\xff\x00\x00\xff'))+png_chunk(b'IEND',b'')
        d=Document();first=d.sections[0]
        first.header.paragraphs[0].text='虚构第一页页眉';first.footer.paragraphs[0].text='虚构第一页页脚'
        first.different_first_page_header_footer=True
        first.first_page_header.paragraphs[0].text='虚构首页页眉';first.first_page_footer.paragraphs[0].text='虚构首页页脚'
        d.add_paragraph('虚构复杂通知',style='Title');d.add_paragraph('致：【商户名称】')
        table=d.add_table(rows=2,cols=3);table.cell(0,0).merge(table.cell(0,2));table.cell(0,0).text='合并表头'
        table.cell(1,0).text='事项';nested=table.cell(1,1).add_table(rows=1,cols=1);nested.cell(0,0).paragraphs[0].add_run('金额：【金额】')
        table.cell(1,2).paragraphs[0].add_run().add_picture(io.BytesIO(png),width=Inches(0.25))
        second=d.add_section(WD_SECTION.NEW_PAGE);second.header.is_linked_to_previous=False;second.footer.is_linked_to_previous=False
        second.header.paragraphs[0].text='虚构第二页页眉';second.footer.paragraphs[0].text='虚构第二页页脚'
        d.add_paragraph('第二页正文与分页设置保留')
        src=self.d/'complex.docx';d.save(src);before=hashlib.sha256(src.read_bytes()).hexdigest()
        out=self.d/'complex-filled.docx';plan=self.d/'complex-plan.json'
        changes=[{'old':'【商户名称】','new':'虚构商户甲'},{'old':'【金额】','new':'900.00'}]
        plan.write_text(json.dumps(changes,ensure_ascii=False),encoding='utf-8')
        self.run_tool('fill_docx.py',src,'--out',out,'--expect-sha256',before,'--replace','【商户名称】=>虚构商户甲','--replace','【金额】=>900.00')
        result=json.loads(self.run_tool('verify_docx_expected.py',src,out,'--plan',plan).stdout)
        self.assertTrue(result['passed']);self.assertTrue(result['same_package_members']);self.assertTrue(result['other_parts_identical']);self.assertTrue(result['body_structure_and_attributes_identical'])
        with zipfile.ZipFile(src) as a,zipfile.ZipFile(out) as b:
            self.assertEqual(set(a.namelist()),set(b.namelist()))
            media=[n for n in a.namelist() if n.startswith('word/media/')]
            self.assertEqual(len(media),1)
            for name in media:self.assertEqual(a.read(name),b.read(name))
        filled=Document(out);self.assertEqual(filled.sections[0].header.paragraphs[0].text,'虚构第一页页眉');self.assertEqual(filled.sections[0].footer.paragraphs[0].text,'虚构第一页页脚')
        self.assertEqual(filled.sections[0].first_page_header.paragraphs[0].text,'虚构首页页眉');self.assertEqual(filled.sections[0].first_page_footer.paragraphs[0].text,'虚构首页页脚')
        self.assertEqual(filled.sections[1].header.paragraphs[0].text,'虚构第二页页眉');self.assertEqual(filled.sections[1].footer.paragraphs[0].text,'虚构第二页页脚')
        table=filled.tables[0];self.assertEqual(table.cell(0,0)._tc,table.cell(0,2)._tc);self.assertEqual(table.cell(1,1).tables[0].cell(0,0).text,'金额：900.00')
        self.assertEqual(len(filled.sections),2);self.assertEqual(len(filled.tables),1)
    def test_install_and_backup_without_overwrite(self):
        import shutil
        source=self.d/'package';(source/'scripts').mkdir(parents=True)
        (source/'SKILL.md').write_text('虚构技能',encoding='utf-8')
        shutil.copy2(SCRIPTS/'install_skill.py',source/'scripts/install_skill.py')
        skills=self.d/'skills';skills.mkdir();backup_root=self.d/'backups';backup_root.mkdir()
        target=skills/'property-work-assistant';env=os.environ.copy();env['PYTHONIOENCODING']='utf-8'
        cmd=[sys.executable,str(source/'scripts/install_skill.py'),'--target',str(target)]
        first=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8',env=env)
        self.assertEqual(first.returncode,0,(first.stdout,first.stderr));self.assertEqual((target/'SKILL.md').read_text(encoding='utf-8'),'虚构技能')
        (target/'local-note.txt').write_text('保留',encoding='utf-8')
        denied=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8',env=env)
        self.assertNotEqual(denied.returncode,0);self.assertEqual((target/'local-note.txt').read_text(encoding='utf-8'),'保留')
        backup=backup_root/'old-version'
        updated=subprocess.run(cmd+['--backup-to',str(backup)],capture_output=True,text=True,encoding='utf-8',env=env)
        self.assertEqual(updated.returncode,0,(updated.stdout,updated.stderr));self.assertEqual((backup/'local-note.txt').read_text(encoding='utf-8'),'保留')
        self.assertFalse((target/'local-note.txt').exists());self.assertEqual((target/'SKILL.md').read_text(encoding='utf-8'),'虚构技能')
        denied2=subprocess.run(cmd+['--backup-to',str(backup)],capture_output=True,text=True,encoding='utf-8',env=env)
        self.assertNotEqual(denied2.returncode,0);self.assertEqual((backup/'local-note.txt').read_text(encoding='utf-8'),'保留')
    def test_installer_preserves_same_machine_preferences_on_update(self):
        import shutil
        profile=self.d/'profile';skills=profile/'skills';skills.mkdir(parents=True)
        pref=profile/'.property-work-assistant'/'preferences.json';pref.parent.mkdir()
        pref_bytes=b'{"schema":1,"preferences":[{"id":"test-only","scope":"demo","preference":"plain table","status":"confirmed-by-user"}]}\n'
        pref.write_bytes(pref_bytes)
        target=skills/'property-work-assistant';backup_root=profile/'skill-backups';backup_root.mkdir()
        env=os.environ.copy();env['PYTHONIOENCODING']='utf-8'
        def package(name,text):
            root=self.d/name/'property-work-assistant';(root/'scripts').mkdir(parents=True)
            (root/'SKILL.md').write_text(text,encoding='utf-8')
            shutil.copy2(SCRIPTS/'install_skill.py',root/'scripts'/'install_skill.py')
            return root/'scripts'/'install_skill.py'
        v1=package('v1','测试技能 v1')
        first=subprocess.run([sys.executable,str(v1),'--target',str(target)],capture_output=True,text=True,encoding='utf-8',env=env)
        self.assertEqual(first.returncode,0,(first.stdout,first.stderr));self.assertEqual(pref.read_bytes(),pref_bytes)
        v2=package('v2','测试技能 v2');backup=backup_root/'before-update'
        update=subprocess.run([sys.executable,str(v2),'--target',str(target),'--backup-to',str(backup)],capture_output=True,text=True,encoding='utf-8',env=env)
        self.assertEqual(update.returncode,0,(update.stdout,update.stderr));self.assertEqual(pref.read_bytes(),pref_bytes)
        self.assertEqual((backup/'SKILL.md').read_text(encoding='utf-8'),'测试技能 v1')
        self.assertEqual((target/'SKILL.md').read_text(encoding='utf-8'),'测试技能 v2')
    def test_installer_restores_old_skill_if_promotion_fails(self):
        import importlib.util,shutil
        from unittest.mock import patch
        source=self.d/'package';(source/'scripts').mkdir(parents=True)
        (source/'SKILL.md').write_text('新技能',encoding='utf-8')
        script=source/'scripts/install_skill.py';shutil.copy2(SCRIPTS/'install_skill.py',script)
        skills=self.d/'skills';skills.mkdir();target=skills/'property-work-assistant';target.mkdir()
        (target/'SKILL.md').write_text('旧技能',encoding='utf-8')
        backup_root=self.d/'backups';backup_root.mkdir();backup=backup_root/'old'
        spec=importlib.util.spec_from_file_location('installer_under_test',script);mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
        original=mod.os.rename
        def rename(src,dst):
            if Path(src).name=='content':raise OSError('模拟安装切换失败')
            return original(src,dst)
        with patch.object(mod.os,'rename',side_effect=rename),patch.object(sys,'argv',['install','--target',str(target),'--backup-to',str(backup)]):
            with self.assertRaises(OSError):mod.main()
        self.assertEqual((target/'SKILL.md').read_text(encoding='utf-8'),'旧技能');self.assertFalse(backup.exists())
    def test_preference_update_and_delete(self):
        store=self.d/'preferences.json'
        r=self.run_tool('office_helper.py','preferences','add','--store',store,'--scope','微信','--text','简洁','--confirmed')
        pid=json.loads(r.stdout)['saved']['id']
        self.run_tool('office_helper.py','preferences','update','--store',store,'--id',pid,'--text','礼貌简短','--confirmed')
        r=self.run_tool('office_helper.py','preferences','list','--store',store);self.assertEqual(json.loads(r.stdout)['preferences'][0]['preference'],'礼貌简短')
        self.run_tool('office_helper.py','preferences','delete','--store',store,'--id',pid,'--confirmed')
        r=self.run_tool('office_helper.py','preferences','list','--store',store);self.assertEqual(json.loads(r.stdout)['preferences'],[])
    def test_discovery_private_path_rejected(self):
        self.run_tool('discover_skills.py',r'C:\private\client',success=False)
    def test_discovery_candidate_url_rejects_dot_segments(self):
        self.assertIsNone(discover_skills.directory_url('vendor/../unsafe'))
        self.assertIsNone(discover_skills.directory_url('vendor/./skill'))
        self.assertIsNone(discover_skills.directory_url('vendor/invalid slug'))
        self.assertEqual(discover_skills.directory_url('vendor/skill-name'),'https://skills.sh/vendor/skill-name')
    def test_discovery_network_failures_return_structured_status(self):
        with patch.object(sys,'argv',['discover_skills.py','pdf ocr']), patch.object(discover_skills.urllib.request,'urlopen',side_effect=http.client.RemoteDisconnected('closed')):
            output=io.StringIO()
            with contextlib.redirect_stdout(output):code=discover_skills.main()
        self.assertEqual(code,3);self.assertEqual(json.loads(output.getvalue())['status'],'unavailable')
        with patch.object(sys,'argv',['discover_skills.py','pdf ocr']), patch.object(discover_skills.urllib.request,'urlopen',return_value=io.BytesIO(b'\xff')):
            output=io.StringIO()
            with contextlib.redirect_stdout(output):code=discover_skills.main()
        self.assertEqual(code,3);self.assertEqual(json.loads(output.getvalue())['status'],'unavailable')
class StrictResult(unittest.TextTestResult):
    def wasSuccessful(self):
        return super().wasSuccessful() and not self.skipped

if __name__=='__main__':
    runner=unittest.main(testRunner=unittest.TextTestRunner(verbosity=2,resultclass=StrictResult),exit=False)
    if runner.result.skipped:
        print('测试不完整：有测试因依赖或条件缺失被跳过，不能判定全部通过。',file=sys.stderr)
        sys.exit(2)
    sys.exit(0 if runner.result.wasSuccessful() else 1)
