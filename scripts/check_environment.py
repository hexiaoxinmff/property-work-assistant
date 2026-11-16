#!/usr/bin/env python3
"""Report local capability prerequisites without installing or sending data."""
import importlib,importlib.metadata,json,os,platform,sys,tempfile
from pathlib import Path

def module(name,package):
    try:
        importlib.import_module(name)
        try:version=importlib.metadata.version(package)
        except importlib.metadata.PackageNotFoundError:version='unknown'
        return {'available':True,'version':version}
    except Exception as exc:return {'available':False,'reason':type(exc).__name__+': '+str(exc)}

def main():
    modules={'Word读写':module('docx','python-docx'),'Excel读写':module('openpyxl','openpyxl'),'文字图片':module('PIL','Pillow')}
    try:
        with tempfile.TemporaryDirectory(prefix='property-assistant-check-') as d:
            p=Path(d)/'probe';p.write_text('ok',encoding='utf-8');assert p.read_text(encoding='utf-8')=='ok'
        temp={'available':True}
    except Exception as exc:temp={'available':False,'reason':type(exc).__name__+': '+str(exc)}
    print(json.dumps({'python':platform.python_version(),'platform':platform.system(),'modules':modules,'temp_writable':temp,
        'limitations':['仅检查模块可导入与临时目录；不证明目标文件、中文字体、网络、图片识别、平台技能加载或跨会话持久化可用。','不会自动安装软件、联网或读取业务资料。']},ensure_ascii=False,indent=2))
if __name__=='__main__':main()
