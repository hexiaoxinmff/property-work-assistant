#!/usr/bin/env python3
"""Download, verify and unpack one pinned public source archive on explicit request."""
from pathlib import Path,PurePosixPath
from urllib.parse import urlparse
import argparse,hashlib,json,shutil,sys,tempfile,urllib.request,zipfile

MAX_DOWNLOAD=100*1024*1024
class CheckedRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self,req,fp,code,msg,headers,newurl):
        p=urlparse(newurl)
        if p.scheme!='https' or p.hostname!='codeload.github.com':raise ValueError('拒绝跳转到非预期下载地址')
        return super().redirect_request(req,fp,code,msg,headers,newurl)

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--name',required=True);p.add_argument('--dest',help='可写的源码目录，默认在技能包下');args=p.parse_args()
    root=Path(__file__).resolve().parent.parent
    entries=json.loads((root/'references'/'upstream-manifest.json').read_text(encoding='utf-8'))
    matches=[e for e in entries if e['name']==args.name]
    if len(matches)!=1:raise ValueError('未找到唯一项目')
    e=matches[0];url=e['archive_url'];u=urlparse(url)
    if u.scheme!='https' or u.hostname!='codeload.github.com' or u.username or u.password:raise ValueError('下载地址不受支持')
    base=Path(args.dest).expanduser().resolve() if args.dest else root/'assets'/'upstream'
    base.mkdir(parents=True,exist_ok=True);target=base/e['name']
    if target.exists():raise FileExistsError('目标已存在，拒绝覆盖')
    with tempfile.TemporaryDirectory(prefix='restore-',dir=base) as tmp:
        tmp=Path(tmp);archive=tmp/'source.zip';h=hashlib.sha256();size=0
        archive_base=(root/'assets'/'upstream-archives').resolve()
        rels=[e['archive_path']] if 'archive_path' in e else e.get('archive_parts',[])
        files=[(root/v).resolve() for v in rels]
        if any(not f.is_relative_to(archive_base) for f in files):raise ValueError('归档路径不在预期目录')
        present=[f.exists() for f in files]
        if any(present) and not all(present):raise ValueError('本地归档不完整，拒绝覆盖或忽略')
        if files and all(present):
            source_kind='local'
            with archive.open('xb') as output:
                for file in files:
                    with file.open('rb') as part:
                        while True:
                            chunk=part.read(1024*1024)
                            if not chunk:break
                            size+=len(chunk)
                            if size>MAX_DOWNLOAD:raise ValueError('归档超过大小上限')
                            h.update(chunk);output.write(chunk)
        else:
            source_kind='network'
            request=urllib.request.Request(url,headers={'User-Agent':'property-office-assistant/0.8'})
            with urllib.request.build_opener(CheckedRedirect()).open(request,timeout=45) as response,archive.open('xb') as output:
                while True:
                    chunk=response.read(1024*1024)
                    if not chunk:break
                    size+=len(chunk)
                    if size>MAX_DOWNLOAD:raise ValueError('下载超过大小上限')
                    h.update(chunk);output.write(chunk)
        if h.hexdigest()!=e['archive_sha256']:raise ValueError('归档哈希不匹配')
        with zipfile.ZipFile(archive) as z:
            if z.testzip():raise ValueError('归档损坏')
            infos=z.infolist();names=[i.filename for i in infos]
            if len(set(names))!=len(names):raise ValueError('归档包含重名路径')
            if sum(i.file_size for i in infos)>200*1024*1024:raise ValueError('解压大小超过限制')
            roots={PurePosixPath(i.filename).parts[0] for i in infos if PurePosixPath(i.filename).parts}
            if len(roots)!=1:raise ValueError('归档根目录异常')
            work=tmp/'contents';work.mkdir()
            for i in infos:
                parts=PurePosixPath(i.filename).parts
                if len(parts)<2 or i.is_dir():continue
                tail=parts[1:]
                if any(x in ('.','..') or '\\' in x or ':' in x for x in tail):raise ValueError('不安全的归档路径')
                if ((i.external_attr>>16)&0o170000)==0o120000:raise ValueError('不允许符号链接')
                file=work.joinpath(*tail);file.parent.mkdir(parents=True,exist_ok=True)
                with z.open(i) as src,file.open('xb') as dst:shutil.copyfileobj(src,dst)
            if target.exists():raise FileExistsError('目标已出现，拒绝覆盖')
            work.rename(target)
    print(json.dumps({'name':e['name'],'sha256':h.hexdigest(),'output':str(target),'source':source_kind},ensure_ascii=False))
if __name__=='__main__':
    try:main()
    except Exception as e:print(json.dumps({'error':str(e)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
