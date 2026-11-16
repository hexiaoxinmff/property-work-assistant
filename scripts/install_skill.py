#!/usr/bin/env python3
"""Install this skill to a user-specified local directory; never download or overwrite."""
from pathlib import Path
import argparse,hashlib,json,os,shutil,sys,tempfile

def digest_tree(root):
    out={}
    for p in root.rglob('*'):
        if '__pycache__' in p.parts or (p.is_file() and p.suffix=='.pyc'):continue
        if p.is_symlink():raise ValueError('技能目录含符号链接，不能自动安装')
        if p.is_file():
            h=hashlib.sha256()
            with p.open('rb') as f:
                for chunk in iter(lambda:f.read(1024*1024),b''):h.update(chunk)
            out[p.relative_to(root).as_posix()]=h.hexdigest()
    return out

def overlaps(a,b):return a==b or a.is_relative_to(b) or b.is_relative_to(a)

def main():
    a=argparse.ArgumentParser(description=__doc__)
    a.add_argument('--target',required=True,help='Full destination path ending in property-work-assistant')
    a.add_argument('--backup-to',help='New backup path outside skills directory, on same filesystem; required for updates')
    args=a.parse_args();source=Path(__file__).resolve().parent.parent
    raw=Path(args.target).expanduser()
    if not raw.is_absolute():raise ValueError('安装路径必须为绝对路径')
    if raw.is_symlink():raise ValueError('安装目标不能是符号链接')
    target=raw.resolve()
    if target.name!='property-work-assistant':raise ValueError('安装目标目录名须为 property-work-assistant')
    if not target.parent.is_dir():raise ValueError('目标技能父目录不存在，请先确认平台支持的目录')
    if overlaps(source,target):raise ValueError('安装目标与源技能包重叠')
    backup=None
    if args.backup_to:
        rawbackup=Path(args.backup_to).expanduser()
        if not rawbackup.is_absolute() or rawbackup.is_symlink():raise ValueError('备份路径须为未存在的绝对路径')
        backup=rawbackup.resolve()
        if backup.exists():raise ValueError('备份路径已存在，拒绝覆盖')
        if not backup.parent.is_dir():raise ValueError('备份父目录不存在')
        if overlaps(backup,target.parent) or overlaps(backup,source):raise ValueError('备份须位于技能目录和源包之外')
    if target.exists():
        if not target.is_dir():raise ValueError('安装目标不是目录')
        if backup is None:raise ValueError('目标已有技能；提供同盘且技能目录之外的全新备份路径后才可更新')
    elif backup is not None:raise ValueError('目标不存在，不需要备份参数')
    source_hashes=digest_tree(source)
    if 'SKILL.md' not in source_hashes:raise ValueError('源技能包缺少 SKILL.md')
    with tempfile.TemporaryDirectory(prefix='.property-install-',dir=target.parent) as tmp:
        stage=Path(tmp)/'content'
        shutil.copytree(source,stage,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
        if digest_tree(stage)!=source_hashes or digest_tree(source)!=source_hashes:raise ValueError('复制或源文件校验失败，未安装')
        moved=False
        try:
            if target.exists():
                try:os.rename(target,backup)
                except OSError as exc:raise ValueError('备份移动失败，请确认备份在同一磁盘且可写：'+str(exc))
                moved=True
            if target.exists():raise FileExistsError('目标在安装期间出现，已停止')
            os.rename(stage,target)
        except BaseException:
            if moved and not target.exists():os.rename(backup,target)
            raise
    print(json.dumps({'target':str(target),'backup':str(backup) if backup else None,'files':len(source_hashes),
        'note':'已核对复制文件；平台是否加载该技能仍需在新会话测试。偏好数据不在技能目录内，未迁移。'},ensure_ascii=False))
if __name__=='__main__':
    try:main()
    except Exception as exc:print(json.dumps({'error':str(exc)},ensure_ascii=False),file=sys.stderr);sys.exit(1)
