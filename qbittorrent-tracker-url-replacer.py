#!/usr/bin/env python3

# 说明：批量替换种子 tracker URL 中的域名/字段（例如站点换域名后保留 passkey 迁移）
#
# 核心安全设计：
# 1. 只替换 URL 中的 host 部分，passkey、路径、查询串原样保留（passkey 丢失会导致做种鉴权失败）
# 2. 只处理「tracker URL 命中旧域名」的种子，其余种子一律不触碰
# 3. 两级执行：--canary 先改 N 个验证，确认无误后再 --apply 批量
# 4. 变更前导出完整映射备份（JSON），可用于审计
# 5. 服务端已保证不会重复添加：新 URL 已存在时 editTracker 返回 Conflict，本脚本如实记录
# 6. editTracker 只会 forceReannounce，不会 recheck/rehash，已下载数据不受影响
#
# 鉴权：支持 API token（>=5.2.0）与 用户名+密码，token 优先
#
# 用法示例：
# 1. 预览：只列出将被修改的种子，不做任何改动
#    python qbittorrent-tracker-url-replacer.py --host http://127.0.0.1:8080 --token qbt_xxx \
#        --old-host t.hddolby.com --new-host www.uhddolby.com
# 2. 小批量试验（默认 2 个），改完立即验证
#    python qbittorrent-tracker-url-replacer.py --host ... --token qbt_xxx \
#        --old-host t.hddolby.com --new-host www.uhddolby.com --canary
# 3. 确认无误后批量执行（带交互确认）
#    python qbittorrent-tracker-url-replacer.py --host ... --token qbt_xxx \
#        --old-host t.hddolby.com --new-host www.uhddolby.com --apply
# 4. 批量执行且跳过确认（仅在已验证 canary 后使用）
#    python qbittorrent-tracker-url-replacer.py --host ... --token qbt_xxx \
#        --old-host t.hddolby.com --new-host www.uhddolby.com --apply --yes

import sys
import os
import json
import time
import argparse
from urllib.parse import urlsplit, urlunsplit

# 保证可从任意工作目录导入同目录下的 qb_auth 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd

from qb_auth import add_auth_args, build_client_from_args

# Windows 控制台默认 GBK 编码，统一改为 UTF-8 并对不可编码字符降级
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# 单个种子 tracker 读取的重试配置
REQUEST_RETRIES = 3
REQUEST_RETRY_DELAY = 1.0

# 服务端 editTracker 的已知冲突/错误返回文本，用于给出可读提示
ERR_TRACKER_NOT_FOUND = 'Tracker not found'
ERR_ALREADY_EXISTS = 'New tracker URL already exists'


def replace_host(url, old_host, new_host):
    """
    仅替换 URL 的 host 字段，其余（scheme / 路径 / 查询串 passkey / 片段）原样保留。
    使用 urlsplit + _replace 避免把 passkey 里的字符误判为分隔结构。
    不匹配 old_host 时返回 None，交由调用方跳过。
    """
    parts = urlsplit(url)
    if parts.hostname is None or parts.hostname.lower() != old_host.lower():
        return None
    # 保留原始端口（若有），只换 host 名
    netloc = new_host
    if parts.port is not None:
        netloc = f'{new_host}:{parts.port}'
    return urlunsplit((parts.scheme, netloc, parts.path, parts.query, parts.fragment))


def fetch_trackers(client, torrent_hash):
    """读取单个种子的 tracker 列表，带有限重试"""
    last_error = None
    for attempt in range(REQUEST_RETRIES):
        try:
            r = client.get('/api/v2/torrents/trackers', params={'hash': torrent_hash})
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last_error = e
            if attempt < REQUEST_RETRIES - 1:
                time.sleep(REQUEST_RETRY_DELAY)
    raise last_error


def build_plan(client, old_host, new_host, progress_step=200):
    """
    扫描全部种子，构建替换计划。
    只纳入「至少有一条 tracker URL 命中旧域名」的种子，其余一律忽略。
    """
    torrents = client.get('/api/v2/torrents/info').json()
    print(f'共 {len(torrents)} 个种子，开始筛选命中 "{old_host}" 的 tracker...')

    # 先用 torrents/info 的 tracker 字段粗筛，避免对 4 万个种子逐个请求 trackers 接口
    candidates = [t for t in torrents if old_host.lower() in (t.get('tracker') or '').lower()]
    print(f'粗筛命中 {len(candidates)} 个种子，逐个读取 tracker 明细...')

    plan = []
    failed = []
    for index, t in enumerate(candidates):
        torrent_hash = t['hash']
        try:
            trackers = fetch_trackers(client, torrent_hash)
        except Exception as e:
            failed.append((t.get('name', ''), torrent_hash, str(e)))
            continue

        for entry in trackers or []:
            old_url = entry.get('url', '')
            # ** 开头为 DHT / PeX / LSD 等伪 tracker，跳过
            if not old_url or old_url.startswith('**'):
                continue
            new_url = replace_host(old_url, old_host, new_host)
            if not new_url or new_url == old_url:
                continue
            plan.append({
                'name': t.get('name', ''),
                'hash': torrent_hash,
                'torrent_state': t.get('state', ''),
                'old_url': old_url,
                'new_url': new_url,
            })

        if progress_step and (index + 1) % progress_step == 0:
            print(f'  已扫描 {index + 1}/{len(candidates)}', flush=True)

    return plan, failed


def apply_plan(client, plan, backup_path):
    """执行替换。逐条调用 editTracker，并记录每条结果。"""
    # 变更前先落盘备份，便于审计与事后核对
    with open(backup_path, 'w', encoding='utf-8') as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)
    print(f'替换计划已备份至: {backup_path}')

    results = []
    for index, item in enumerate(plan, 1):
        r = client.post('/api/v2/torrents/editTracker',
                        data={'hash': item['hash'],
                              'url': item['old_url'],
                              'newUrl': item['new_url']})
        body = (r.text or '').strip()
        # 与登录接口一致：新版 qb 对 editTracker 成功返回 204 且响应体为空，故用 2xx 判定
        if 200 <= r.status_code < 300:
            status, note = '成功', ''
        elif r.status_code == 409 or ERR_ALREADY_EXISTS in body:
            status, note = '跳过', '新 URL 已存在'  # 幂等：目标已是期望状态
        elif ERR_TRACKER_NOT_FOUND in body:
            status, note = '失败', '旧 tracker 未找到（可能已被修改）'
        else:
            status, note = '失败', f'HTTP {r.status_code} {body[:120]}'

        results.append({**item, 'status': status, 'note': note})

        if status == '失败':
            print(f'  [失败] {item["name"][:40]}: {note}')
        if index % 100 == 0:
            print(f'  已处理 {index}/{len(plan)}', flush=True)

    return results


def verify(client, results):
    """改后复核：重新读取 tracker，确认新 URL 已生效且旧 URL 已消失"""
    ok = skipped = bad = 0
    for item in results:
        if item['status'] == '失败':
            bad += 1
            continue
        try:
            trackers = fetch_trackers(client, item['hash'])
        except Exception:
            bad += 1
            continue
        urls = [e.get('url', '') for e in trackers or []]
        if item['new_url'] in urls and item['old_url'] not in urls:
            ok += 1
        elif item['new_url'] in urls:
            skipped += 1  # 新 URL 在，但旧 URL 仍在（多 tracker 场景）
        else:
            bad += 1
    return ok, skipped, bad


def export_excel(results, out_path):
    df = pd.DataFrame(results)
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='替换明细', index=False)
    return out_path


def main():
    p = argparse.ArgumentParser(
        description='批量替换种子 tracker URL 的域名（保留 passkey 等其余字段）')
    p.add_argument('--host', '-H', required=True, help='qBittorrent Web UI 地址')
    add_auth_args(p)
    p.add_argument('--old-host', required=True, help='旧域名，例如 t.hddolby.com')
    p.add_argument('--new-host', required=True, help='新域名，例如 www.uhddolby.com')
    p.add_argument('--canary', action='store_true', help='小批量试验模式：只改前 N 个并验证')
    p.add_argument('--canary-count', type=int, default=2, help='试验数量，默认 2')
    p.add_argument('--apply', action='store_true', help='批量执行替换')
    p.add_argument('--yes', action='store_true', help='跳过批量执行前的交互确认')
    p.add_argument('--backup', default='tracker_replace_backup.json', help='变更前计划备份(JSON)')
    p.add_argument('--out', '-o', default='tracker_replace_result.xlsx', help='结果 Excel 文件名')
    p.add_argument('--progress-step', type=int, default=200, help='扫描进度打印间隔')

    args = p.parse_args()
    if args.canary and args.apply:
        p.error('--canary 与 --apply 互斥：请先 --canary 验证，再 --apply 批量')

    client = build_client_from_args(args)
    print(f'鉴权方式: {"API token" if client.auth_mode == "token" else "用户名密码(Cookie 会话)"}')

    try:
        run(client, args)
    finally:
        client.logout()


def run(client, args):
    plan, failed = build_plan(client, args.old_host, args.new_host, args.progress_step)

    if failed:
        print(f'\n[警告] {len(failed)} 个种子 tracker 读取失败，已跳过：')
        for name, torrent_hash, err in failed[:5]:
            print(f'  - {name[:40]} ({torrent_hash[:12]}): {err}')

    if not plan:
        print(f'\n未发现任何包含 "{args.old_host}" 的 tracker，无需修改。')
        return

    # 同一 URL 可能被多个种子共用，去重后展示，便于确认替换规则
    unique_pairs = sorted({(i['old_url'], i['new_url']) for i in plan})
    affected = len({i['hash'] for i in plan})

    print(f'\n替换规则: {args.old_host} -> {args.new_host}')
    print(f'涉及种子 {affected} 个，待替换 tracker 条目 {len(plan)} 条')
    print(f'去重后的 URL 变更 {len(unique_pairs)} 组，示例：')
    for old_url, new_url in unique_pairs[:3]:
        # 打印时截断 passkey，避免完整密钥落入终端与日志
        print(f'  {old_url[:72]}...')
        print(f'  -> {new_url[:72]}...')

    if not (args.canary or args.apply):
        print('\n预览模式：未对 qBittorrent 做任何修改。')
        print('  先执行 --canary 小批量试验，确认无误后再 --apply 批量执行。')
        return

    if args.canary:
        target = plan[:args.canary_count]
        print(f'\n[试验模式] 仅修改前 {len(target)} 条：')
        for item in target:
            print(f'  - {item["name"][:50]}')
        results = apply_plan(client, target, args.backup)
    else:
        print(f'\n[批量模式] 将修改 {len(plan)} 条 tracker，涉及 {affected} 个种子。')
        if not args.yes:
            answer = input('确认执行？请输入 yes 继续: ').strip().lower()
            if answer != 'yes':
                print('已取消，未做任何修改。')
                return
        results = apply_plan(client, plan, args.backup)

    ok = len([r for r in results if r['status'] == '成功'])
    skipped = len([r for r in results if r['status'] == '跳过'])
    bad = len([r for r in results if r['status'] == '失败'])
    print(f'\n执行结果: 成功 {ok}，跳过 {skipped}，失败 {bad}')

    print('正在复核...')
    v_ok, v_partial, v_bad = verify(client, [r for r in results if r['status'] != '跳过'])
    print(f'复核结果: 新 URL 已生效 {v_ok}，旧 URL 仍残留 {v_partial}，异常 {v_bad}')

    export_excel(results, args.out)
    print(f'明细已导出: {args.out}')

    if args.canary:
        print('\n请先在 qBittorrent 中确认这几个种子状态正常（tracker 已连上、未触发校验）。')
        print('确认无误后，执行 --apply 进行批量替换。')


if __name__ == '__main__':
    main()
