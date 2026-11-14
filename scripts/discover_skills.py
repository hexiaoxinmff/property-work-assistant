#!/usr/bin/env python3
"""Search public skills.sh directory with a generic, non-sensitive task keyword."""
import argparse,http.client,json,re,sys,urllib.error,urllib.parse,urllib.request

def clean(value,max_len=180):
    value=str(value or '')
    return ''.join(c for c in value if c.isprintable())[:max_len]

def directory_url(slug):
    parts=slug.split('/')
    if len(parts)<2 or len(parts)>8 or any(part in ('.','..') or not re.fullmatch(r'[A-Za-z0-9._-]{1,100}',part) for part in parts):
        return None
    return 'https://skills.sh/'+urllib.parse.quote(slug,safe='/')

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('query',help='Generic capability words only; never include private task details')
    p.add_argument('--limit',type=int,default=8)
    args=p.parse_args()
    query=' '.join(args.query.split())
    if not 2<=len(query)<=80:raise ValueError('查询词须为 2 至 80 个字符')
    if not 1<=args.limit<=20:raise ValueError('limit 须为 1 至 20')
    if re.search(r'@|https?://|[\\/]|[0-9]{6,}',query,re.I):
        raise ValueError('查询词应使用不含邮箱、路径、链接或长数字的通用任务关键词')
    endpoint='https://skills.sh/api/search?'+urllib.parse.urlencode({'q':query})
    request=urllib.request.Request(endpoint,headers={'User-Agent':'property-assistant-skill-discovery/0.1','Accept':'application/json'})
    try:
        with urllib.request.urlopen(request,timeout=12) as response:
            raw=response.read(1024*1024+1)
        if len(raw)>1024*1024:raise ValueError('搜索响应过大')
        data=json.loads(raw)
        if not isinstance(data,dict) or not isinstance(data.get('skills'),list):raise ValueError('搜索响应格式不符合预期')
        results=[]
        for item in data['skills'][:args.limit]:
            if not isinstance(item,dict):continue
            slug=clean(item.get('id'));source=clean(item.get('source'));name=clean(item.get('name'))
            url=directory_url(slug)
            results.append({'name':name,'source':source,'directory_url':url,'installs':item.get('installs') if isinstance(item.get('installs'),int) else None})
        result={'status':'ok' if results else 'no_results','query':query,'results':results,'note':'搜索结果为第三方元数据；需核对来源、内容、权限与实际适用性。'}
    except (urllib.error.URLError,http.client.HTTPException,TimeoutError,OSError,UnicodeError,ValueError,json.JSONDecodeError) as exc:
        result={'status':'unavailable','reason':clean(type(exc).__name__+': '+str(exc)), 'results':[], 'note':'搜索不可用，继续使用现有技能或平台能力。'}
    print(json.dumps(result,ensure_ascii=False,indent=2));return 0 if result['status']!='unavailable' else 3
if __name__=='__main__':
    try:sys.exit(main())
    except Exception as exc:
        print(json.dumps({'status':'invalid_request','reason':str(exc),'results':[]},ensure_ascii=False));sys.exit(2)
