#!/usr/bin/env python3

# 说明：将指定状态种子改为做种状态（仅在Windows上测试过）
# 主要功能：
# 1. 导出指定状态的种子信息到Excel
# 2. 跳过校验中的种子（将其改为做种状态）
# 3. 解除强制做种状态（改为普通做种状态）

# 鉴权：支持 API token（qBittorrent >= 5.2.0，推荐）与 用户名+密码，同时提供时 token 优先
#
# 用法示例（用 --token 或 --username/--password 均可）：
# 1. 将校验状态种子改为做种状态（预览模式）：
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter checking --out checking.xlsx
# 2. 实际执行跳过校验（确认无误后加上--skip-check）：
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter checking --out checking.xlsx --skip-check
# 3. 解除强制做种状态（预览模式）
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter forced --out forced.xlsx
# 4. 实际执行解除强制做种（确认无误后加上--skip-forced）：
#    python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter forced --out forced.xlsx --skip-forced

# 兼容性说明（qBittorrent 5.3.0 / WebAPI v2.16）:
# 1. 跳过校验的参数由 skip_checking 改名为 seedMode；服务端会静默忽略不认识的参数，
#    所以两个名字一起发，任何一端都只解析自己认识的那个。
# 2. 添加后是否启动的参数由 paused 改名为 stopped，同样双写并显式置 false，
#    否则新种子会跟随全局偏好被加成 stoppedUP/stoppedDL（即"暂停状态"）。
# 3. torrents/delete 只把种子摘出会话，libtorrent 真正移除完成后才能同 hash 重加，
#    删完立刻重加会被 duplicate_is_error 静默丢弃（HTTP 仍返回成功），故先轮询确认删除完成。
# 4. torrents/add 成功可能返回 202（pending_count > 0），一条都没加进去时返回 409，
#    判定一律用 2xx 并解析 JSON 里的 success/failure/pending 计数，不能写死 200。

# 过滤选项: checking, completed, forced, all

import os
import sys

# 保证可从任意工作目录导入同目录下的 qb_auth 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse
import json
import re
import time

import pandas as pd
import requests

from qb_auth import add_auth_args, build_client_from_args

# Windows 控制台默认 GBK 编码，无法输出 emoji 等字符，统一改为 UTF-8 并对不可编码字符降级
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

def fetch_torrents(client):
    """拉取全部种子信息。41015 个种子时响应可达数十 MB，读取超时已放宽至 60s。"""
    r = client.get('/api/v2/torrents/info')
    r.raise_for_status()
    return r.json()

CHECKING_STATES = ('checkingDL', 'checkingUP')
# 跳过校验成功后应处于做种态；5.x 的停止态叫 stoppedUP，旧版仍叫 pausedUP
DONE_STATES = ('uploading', 'stalledUP', 'queuedUP', 'forcedUP', 'stoppedUP', 'pausedUP')
BAD_STATES = ('missingFiles', 'error', 'metaDL', 'forcedMetaDL')

POLL_INTERVAL = 0.5       # 轮询间隔
DELETE_WAIT_TIMEOUT = 30  # 等待删除真正完成的上限
ADD_WAIT_TIMEOUT = 30     # 等待种子重加并进入做种态的上限
ADD_RETRY_LIMIT = 3       # 重加未能落地时最多尝试几次
RETRY_PAUSE = 2           # 每次重试前的等待，给 libtorrent 完成上一轮移除留出时间


def is_ok(response):
    """成功判定用 2xx：5.3.0 的 torrents/add 在 pending 时返回 202，写死 200 会把成功当失败"""
    return 200 <= response.status_code < 300


def find_torrent(client, torrent_hash):
    """用 hashes 参数精确查询单个种子，避免每处理一个就拉一次全量（本库 4 万多种子）"""
    response = client.get('/api/v2/torrents/info', params={'hashes': torrent_hash})
    if not is_ok(response):
        raise RuntimeError(f'查询种子状态失败: HTTP {response.status_code} {response.text[:200]}')
    torrents = response.json()
    return torrents[0] if torrents else None


def wait_until_removed(client, torrent_hash, timeout=DELETE_WAIT_TIMEOUT):
    """等待种子从会话中彻底消失——libtorrent 移除是异步的，没消失就重加必然被静默丢弃"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if find_torrent(client, torrent_hash) is None:
            return True
        time.sleep(POLL_INTERVAL)
    return False


def bool_str(value):
    return 'true' if value else 'false'


def normalize_tags(tags):
    """info 返回的 tags 形如 "a, b"，add 需要不含空格的逗号分隔串"""
    if not tags:
        return ''
    return ','.join(part.strip() for part in tags.split(',') if part.strip())


def build_add_params(torrent):
    """
    组装重加参数：跳过校验 + 回填原有属性。

    AutoTMM 打开时不能发 savepath——服务端看到 savepath 会隐式关掉 AutoTMM，
    路径改由分类/默认规则推导，所以此时只声明 autoTMM=true。
    forced 故意不回填：解除强制做种正是 --skip-forced 的目的。
    """
    params = {'skip_checking': 'true', 'seedMode': 'true'}
    params.update({'paused': 'false', 'stopped': 'false'})

    if torrent.get('auto_tmm'):
        params['autoTMM'] = 'true'
    else:
        params['autoTMM'] = 'false'
        save_path = torrent.get('save_path') or ''
        if save_path:
            params['savepath'] = save_path

    if torrent.get('category'):
        params['category'] = torrent['category']
    tags = normalize_tags(torrent.get('tags', ''))
    if tags:
        params['tags'] = tags
    if torrent.get('name'):
        params['rename'] = torrent['name']

    params['upLimit'] = torrent.get('up_limit', -1)
    params['dlLimit'] = torrent.get('dl_limit', -1)
    params['sequentialDownload'] = bool_str(torrent.get('seq_dl'))
    params['firstLastPiecePrio'] = bool_str(torrent.get('f_l_piece_prio'))
    params['ratioLimit'] = torrent.get('ratio_limit', -2)
    params['seedingTimeLimit'] = torrent.get('seeding_time_limit', -2)
    params['inactiveSeedingTimeLimit'] = torrent.get('inactive_seeding_time_limit', -2)
    if torrent.get('share_limits_mode'):
        params['shareLimitsMode'] = torrent['share_limits_mode']
    if torrent.get('share_limit_action'):
        params['shareLimitAction'] = torrent['share_limit_action']
    return params


def parse_add_reply(response):
    """5.3.0 起 add 返回 JSON 计数，旧版返回 "Ok."/"Err." 文本，两种格式都要判"""
    body = response.text.strip()
    if not body.startswith('{'):
        if body.startswith('Err'):
            return False, f'服务端拒绝: {body[:200]}'
        return True, 'ok'

    try:
        result = json.loads(body)
    except ValueError:
        return True, 'ok'

    added = result.get('success_count', 0)
    pending = result.get('pending_count', 0)
    if not added and not pending:
        return False, f"未添加任何种子（failure_count={result.get('failure_count', 0)}）"
    return True, f'success={added} pending={pending}'


def reimport_with_skip_check(client, torrent, torrent_file):
    """重新导入种子（跳过校验），返回 (是否受理, 说明)"""
    files = {'torrents': (f"{torrent['name']}.torrent", torrent_file, 'application/x-bittorrent')}
    response = client.post('/api/v2/torrents/add', data=build_add_params(torrent), files=files)

    if response.status_code == 409:
        return False, '重加被拒绝（409 Conflict）：同 hash 种子仍在会话中，删除尚未完成'
    if not is_ok(response):
        return False, f'HTTP {response.status_code} {response.text[:200]}'
    return parse_add_reply(response)


def verify_skipped(client, torrent_hash, timeout=ADD_WAIT_TIMEOUT):
    """
    确认种子已回到会话并进入做种态，返回 (结论, 说明)。

    结论取值：ok=已做种；suspect=仍在慢慢校验（说明跳过并未真正生效）；
    missing=会话里查不到（这次 add 被静默丢弃，可重试）；failed=状态异常。
    只认明确的完成态（DONE_STATES），不能改用"非校验态 + 进度 100%"判定：
    刚重加时服务端会短暂返回 checkingResumeData，此时进度同样已经是 1，容易被误判成跳过成功。
    """
    deadline = time.time() + timeout
    seen = False
    while time.time() < deadline:
        current = find_torrent(client, torrent_hash)
        if current is not None:
            seen = True
            state = current.get('state', '')
            if state in BAD_STATES:
                return 'failed', f'状态异常: {state}'
            if state in DONE_STATES:
                return 'ok', state
        time.sleep(POLL_INTERVAL)

    if not seen:
        return 'missing', '重加后未出现在会话中'
    return 'suspect', f'等待超时，仍在校验: {current.get("state", "")}'


def reimport_and_confirm(client, torrent, torrent_file):
    """
    重新导入并确认结果，返回 (结论, 说明)。

    会话层删除是同步的，libtorrent 层的原生句柄移除却是异步的：抢先 add 会撞上
    duplicate_is_error 而被静默丢弃（上游 #24367），因此对 missing 与 409 都要退避重试。
    """
    torrent_hash = torrent['hash']
    detail = ''
    for attempt in range(1, ADD_RETRY_LIMIT + 1):
        accepted, accepted_detail = reimport_with_skip_check(client, torrent, torrent_file)
        detail = accepted_detail

        if accepted:
            verdict, state_detail = verify_skipped(client, torrent_hash)
            if verdict != 'missing':
                return verdict, f'{detail} / {state_detail}'
            detail = state_detail
        elif not accepted_detail.startswith('重加被拒绝（409'):
            # 非 409 的失败与删除未完成同类，重试没有意义
            return 'failed', detail

        if attempt < ADD_RETRY_LIMIT:
            print(f'♻️ 第 {attempt} 次重加未确认成功（{detail}），{RETRY_PAUSE}s 后重试')
            time.sleep(RETRY_PAUSE)

    return 'suspect', f'{ADD_RETRY_LIMIT} 次重加后仍无法在会话中确认: {detail}'


def save_torrent_backup(backup_dir, torrent_name, torrent_hash, torrent_file):
    """
    把导出的 .torrent 落盘。

    tracker 的 passkey 只存在于种子文件里，任务记录一旦丢了无法凭记忆重建，
    所以重加失败时必须留下可手动导入的原件。
    """
    os.makedirs(backup_dir, exist_ok=True)
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', torrent_name)[:80]
    backup_path = os.path.join(backup_dir, f'{safe_name}-{torrent_hash[:8]}.torrent')
    with open(backup_path, 'wb') as handle:
        handle.write(torrent_file)
    return backup_path


def skip_reimport(client, torrents, backup_dir):
    """
    删除并重新导入种子，用 seedMode（旧版为 skip_checking）跳过校验。

    返回 {'ok': [...], 'suspect': [...], 'failed': [...]}。suspect 指重试后仍未确认
    （还在校验或压根没落地），failed 指明确失败或状态异常；两者都已备份 .torrent。
    """
    results = {'ok': [], 'suspect': [], 'failed': []}
    total = len(torrents)

    for index, torrent in enumerate(torrents, 1):
        torrent_hash = torrent['hash']
        torrent_name = torrent['name']
        tag = f'[{index}/{total}]'

        # 导出种子文件（含 tracker 与 passkey，比从种子目录复制可靠）
        # 删除前必须确认导出可用：bencode 字典以 'd' 开头，否则一旦删完再重加就无路可退
        response = client.get('/api/v2/torrents/export', params={'hash': torrent_hash})
        if not is_ok(response) or not response.content or not response.content.startswith(b'd'):
            print(f'❌ {tag} 获取种子文件失败，未做任何改动: {torrent_name} (HTTP {response.status_code})')
            results['failed'].append(torrent)
            continue
        torrent_file = response.content

        # 删除种子（保留文件）
        response = client.post('/api/v2/torrents/delete',
                               data={'hashes': torrent_hash, 'deleteFiles': 'false'})
        if not is_ok(response):
            print(f'❌ {tag} 删除种子失败 {torrent_name}: HTTP {response.status_code} {response.text[:200]}')
            results['failed'].append(torrent)
            continue

        # 等删除完成后再重加，否则会撞上同 hash 的 duplicate 而被静默丢弃
        if not wait_until_removed(client, torrent_hash):
            backup_path = save_torrent_backup(backup_dir, torrent_name, torrent_hash, torrent_file)
            print(f'❌ {tag} 删除在 {DELETE_WAIT_TIMEOUT}s 内未完成，暂不重加: {torrent_name}')
            print(f'    种子文件已备份: {backup_path}')
            results['failed'].append(torrent)
            continue

        verdict, detail = reimport_and_confirm(client, torrent, torrent_file)
        if verdict == 'ok':
            print(f'✅ {tag} 已跳过校验: {torrent_name} → {detail}')
            results['ok'].append(torrent)
            continue

        # 只要没确认成功就留档：此刻原记录已删除，而新记录可能压根没落地
        backup_path = save_torrent_backup(backup_dir, torrent_name, torrent_hash, torrent_file)
        if verdict == 'suspect':
            print(f'⚠️ {tag} 结果存疑，需人工确认: {torrent_name}（{detail}）')
            results['suspect'].append(torrent)
        else:
            print(f'❌ {tag} 重加未成功: {torrent_name}（{detail}）')
            results['failed'].append(torrent)
        print(f'    种子文件已备份，可手动重新导入: {backup_path}')

    return results


def report_skip_result(title, results):
    """打印单个批次的处理结果汇总"""
    print(f'{title}: 成功 {len(results["ok"])} / 存疑 {len(results["suspect"])} / 失败 {len(results["failed"])}')

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

def print_server_info(client):
    """打印服务端版本，跳过校验的可用参数名随版本而变，出问题时先核对这一行"""
    try:
        qb_version = client.get('/api/v2/app/version').text.strip()
        api_version = client.get('/api/v2/app/webapiVersion').text.strip()
        print(f'服务端: qBittorrent {qb_version} / WebAPI v{api_version}')
    except requests.RequestException as error:
        print(f'服务端版本获取失败（不影响执行）: {error}')

def main():
    p = argparse.ArgumentParser()
    p.add_argument('--host','-H', required=True, help='qBittorrent Web UI 地址，例如 http://127.0.0.1:8080')
    add_auth_args(p)
    p.add_argument('--filter','-f',
                   choices=['checking','completed','forced','all'],
                   default='checking',
                   help='导出过滤: checking=校验中, completed=已完成, forced=强制做种, all=全部'
                  )
    p.add_argument('--out','-o', default='qb_torrents.xlsx', help='输出Excel文件名')
    p.add_argument('--skip-check', action='store_true', help='跳过所有正在校验的种子')
    p.add_argument('--skip-forced', action='store_true', help='跳过所有强制做种的种子')
    p.add_argument('--backup-dir', default='skip_check_backup',
                   help='重加失败时存放导出 .torrent 的目录（仅在确实失败时才创建）')
    
    args = p.parse_args()
    client = build_client_from_args(args)
    print(f'鉴权方式: {"API token" if client.auth_mode == "token" else "用户名密码(Cookie 会话)"}')
    print_server_info(client)

    try:
        run(client, args)
    finally:
        # Cookie 模式下主动登出，避免服务端会话堆积；token 模式无状态，无需处理
        client.logout()

def run(client, args):
    torrents = fetch_torrents(client)

    # 跳过校验中的种子
    if args.skip_check:
        checking_torrents = [t for t in torrents if t.get('state') in CHECKING_STATES]
        print(f'待跳过校验: {len(checking_torrents)} 个')
        if checking_torrents:
            report_skip_result('跳过校验', skip_reimport(client, checking_torrents, args.backup_dir))
        torrents = fetch_torrents(client)  # 重新获取最新状态

    # 跳过强制做种的种子
    if args.skip_forced:
        forced_torrents = [t for t in torrents if t.get('state') == 'forcedUP']
        print(f'待解除强制做种: {len(forced_torrents)} 个')
        if forced_torrents:
            report_skip_result('解除强制做种', skip_reimport(client, forced_torrents, args.backup_dir))
        torrents = fetch_torrents(client)  # 重新获取最新状态
    
    # 导出逻辑
    rows = []
    for t in torrents:
        state = t.get('state','')
        progress = float(t.get('progress',0.0))
        
        if args.filter == 'checking' and state not in CHECKING_STATES:
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