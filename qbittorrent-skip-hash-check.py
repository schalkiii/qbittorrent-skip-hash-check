#!/usr/bin/env python3

# 说明：将指定状态种子改为做种状态（仅在Windows上测试过）
# 主要功能：
# 1. 导出指定状态的种子信息到Excel
# 2. 跳过校验中的种子（将其改为做种状态）
# 3. 解除强制做种状态（改为普通做种状态）

# 用法示例：
# 1. 将校验状态种子改为做种状态（预览模式）：
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password YOUR_PASSWORD --filter checking --out checking.xlsx
# 2. 实际执行跳过校验（确认无误后加上--skip-check）：
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password YOUR_PASSWORD --filter checking --out checking.xlsx --skip-check
# 3. 解除强制做种状态（预览模式）
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password YOUR_PASSWORD --filter forced --out forced.xlsx
# 4. 实际执行解除强制做种（确认无误后加上--skip-forced）：
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password YOUR_PASSWORD --filter forced --out forced.xlsx --skip-forced

# 过滤选项: checking, completed, forced, all

import requests
import argparse
import sys
import time
import pandas as pd

def login(session, base_url, username, password):
    url = base_url.rstrip('/') + '/api/v2/auth/login'
    r = session.post(url, data={'username': username, 'password': password})
    if r.status_code != 200 or 'Ok.' not in r.text:
        raise RuntimeError(f'登录失败: {r.status_code} {r.text}')
    return True

def fetch_torrents(session, base_url):
    url = base_url.rstrip('/') + '/api/v2/torrents/info'
    r = session.get(url)
    r.raise_for_status()
    return r.json()

def skip_reimport(session, base_url, torrents):
    """
    删除并重新导入种子，利用 skip_checking 跳过校验
    """
    count = 0
    for t in torrents:
        torrent_hash = t['hash']
        save_path = t['save_path']
        torrent_name = t['name']
        
        # 导出种子文件
        dl_url = base_url.rstrip('/') + '/api/v2/torrents/export'
        r = session.get(dl_url, params={'hash': torrent_hash})
        if r.status_code != 200:
            print(f"❌ 获取种子失败 {torrent_name}: {r.text}")
            continue
        torrent_file = r.content
        
        # 删除种子 (不删文件)
        del_url = base_url.rstrip('/') + '/api/v2/torrents/delete'
        r = session.post(del_url, data={'hashes': torrent_hash, 'deleteFiles': 'false'})
        if r.status_code != 200:
            print(f"❌ 删除种子失败 {torrent_name}: {r.text}")
            continue
        
        # 重新添加，跳过校验
        add_url = base_url.rstrip('/') + '/api/v2/torrents/add'
        files = {'torrents': (f"{torrent_name}.torrent", torrent_file, 'application/x-bittorrent')}
        data = {'savepath': save_path, 'skip_checking': 'true'}
        r = session.post(add_url, data=data, files=files)
        if r.status_code != 200:
            print(f"❌ 重新添加失败 {torrent_name}: {r.text}")
            continue
        
        print(f"✅ 已跳过校验并重新导入: {torrent_name}")
        count += 1
    return count

def pick_trackers(t):
    if t.get('tracker') and isinstance(t.get('tracker'), str):
        return t.get('tracker')
    if t.get('tracker_names'):
        return ','.join(t.get('tracker_names'))
    if t.get('trackers'):
        try:
            return ','.join([tt.get('url','') for tt in t.get('trackers')])
        except:
            pass
    return ''

def nice_bytes(n):
    try:
        n = int(n)
    except:
        return str(n)
    for unit in ['B','KB','MB','GB','TB','PB']:
        if n < 1024:
            return f"{n:.2f}{unit}"
        n /= 1024
    return f"{n:.2f}EB"

def main():
    p = argparse.ArgumentParser()
    p.add_argument('--host','-H', required=True, help='qBittorrent Web UI 地址，例如 http://127.0.0.1:8080')
    p.add_argument('--username','-u', required=True)
    p.add_argument('--password','-p', required=True)
    p.add_argument('--filter','-f',
                   choices=['checking','completed','forced','all'],
                   default='checking',
                   help='导出过滤: checking=校验中, completed=已完成, forced=强制做种, all=全部'
                  )
    p.add_argument('--out','-o', default='qb_torrents.xlsx', help='输出Excel文件名')
    p.add_argument('--skip-check', action='store_true', help='跳过所有正在校验的种子')
    p.add_argument('--skip-forced', action='store_true', help='跳过所有强制做种的种子')
    
    args = p.parse_args()
    s = requests.Session()
    
    try:
        login(s, args.host, args.username, args.password)
    except Exception as e:
        print('登录失败:', e)
        sys.exit(1)
    
    torrents = fetch_torrents(s, args.host)
    
    # 跳过校验中的种子
    if args.skip_check:
        checking_torrents = [t for t in torrents if t.get('state') in ('checkingDL','checkingUP')]
        count = skip_reimport(s, args.host, checking_torrents)
        print(f'已跳过 {count} 个校验中的种子')
        torrents = fetch_torrents(s, args.host)  # 重新获取最新状态
    
    # 跳过强制做种的种子
    if args.skip_forced:
        forced_torrents = [t for t in torrents if t.get('state') == 'forcedUP']
        count = skip_reimport(s, args.host, forced_torrents)
        print(f'已跳过 {count} 个强制做种的种子')
        torrents = fetch_torrents(s, args.host)  # 重新获取最新状态
    
    # 导出逻辑
    rows = []
    for t in torrents:
        state = t.get('state','')
        progress = float(t.get('progress',0.0))
        
        if args.filter == 'checking':
            if state not in ('checkingDL','checkingUP'):
                continue
        elif args.filter == 'completed':
            if progress < 0.999999:
                continue
        elif args.filter == 'forced':
            if state != 'forcedUP':
                continue
        
        rows.append({
            'name': t.get('name'),
            'hash': t.get('hash'),
            'state': state,
            'progress': round(progress*100,4),
            'size_bytes': t.get('size', t.get('total_size', '')),
            'size_readable': nice_bytes(t.get('size', t.get('total_size', 0))),
            'save_path': t.get('save_path', t.get('content_path', '')),
            'num_files': t.get('num_files', t.get('filecount', '')),
            'completed_pieces': t.get('completed_pieces',''),
            'ratio': t.get('ratio',''),
            'seeds': t.get('num_seeds', t.get('num_complete', '')),
            'peers': t.get('num_connections', t.get('num_incomplete', '')),
            'added_on': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t.get('added_on',0))) if t.get('added_on') else '',
            'last_activity': time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t.get('last_activity',0))) if t.get('last_activity') else '',
            'category': t.get('category',''),
            'trackers': pick_trackers(t),
            'label': t.get('label',''),
            'files': (', '.join([f.get('name','') for f in t.get('files',[])]) if isinstance(t.get('files',[]), list) and t.get('files') else '')
        })
    
    if not rows:
        print('没有匹配的种子。')
        return
    
    df = pd.DataFrame(rows)
    df.to_excel(args.out, index=False)
    print(f'导出完成: {args.out} (共 {len(rows)} 条)')

if __name__ == '__main__':
    main()