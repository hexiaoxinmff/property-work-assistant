#!/usr/bin/env python3
"""Verify and unpack one bundled third-party source archive without overwriting."""
from pathlib import Path,PurePosixPath
import argparse,hashlib,json,shutil,sys,tempfile,zipfile

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--name',required=True);args=p.parse_args()
    root=Path(__file__).resolve().parent.parent
    entries=json.loads((root/'references'/'upstream-manifest.json').read_text(encoding='utf-8'))
    matches=[e for e in entries if e['name']==args.name]
    if len(matches)!=1:raise ValueError('未找到唯一的项目')
    e=matches[0];archive_root=(root/'assets'/'upstream-archives').resolve()
    source_paths=[e['archive_path']] if 'archive_path' in e else e.get('archive_parts',[])
    if not source_paths:raise ValueError('归档缺少文件或分片')
    target_base=root/'assets'/'upstream';target_base.mkdir(parents=True,exist_ok=True)
    target=target_base/e['name']
    if target.exists():raise FileExistsError('目标已存在，不覆盖：'+str(target))
    with tempfile.TemporaryDirectory(prefix='restore-',dir=target_base) as temp:
        temp_root=Path(temp);assembled=temp_root/'archive.zip';digest=hashlib.sha256()
        with assembled.open('xb') as dest:
            for relative in source_paths:
                source=(root/relative).resolve()
                if not source.is_relative_to(archive_root):raise ValueError('归档路径不在允许目录')
                with source.open('rb') as f:
                    for block in iter(lambda:f.read(1024*1024),b''):
                        digest.update(block);dest.write(block)
        if digest.hexdigest()!=e['archive_sha256']:raise ValueError('归档哈希不匹配')
        with zipfile.ZipFile(assembled) as z:
            if z.testzip():raise ValueError('归档损坏')
            infos=z.infolist();names=[i.filename for i in infos]
            if len(set(names))!=len(names):raise ValueError('归档包含重名路径')
            if sum(i.file_size for i in infos)>200*1024*1024:raise ValueError('解压大小超出限制')
            roots={PurePosixPath(i.filename).parts[0] for i in infos if PurePosixPath(i.filename).parts}
            if len(roots)!=1:raise ValueError('归档根目录异常')
            work=temp_root/'contents';work.mkdir()
            for i in infos:
                parts=PurePosixPath(i.filename).parts
                if len(parts)<2 or i.is_dir():continue
                tail=parts[1:]
                if any(p in ('.','..') or '\\' in p or ':' in p for p in tail):raise ValueError('不安全的归档路径')
                if ((i.external_attr>>16)&0o170000)==0o120000:raise ValueError('不允许符号链接')
                f=work.joinpath(*tail);f.parent.mkdir(parents=True,exist_ok=True)
                with z.open(i) as src,f.open('xb') as dst:shutil.copyfileobj(src,dst)
            if target.exists():raise FileExistsError('目标已出现，拒绝覆盖')
            work.rename(target)
    print(json.dumps({'name':e['name'],'sha256':e['archive_sha256'],'output':str(target)},ensure_ascii=False,indent=2))
if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
