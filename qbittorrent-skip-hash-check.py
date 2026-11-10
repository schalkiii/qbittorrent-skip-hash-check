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
# 5. 批量流水线：每批先全部 add 完，再整批收敛复核（先添加、后检查），
#    避免 N 个 0.5s 等待窗口串行累加；状态查询用 hashes 分片，绝不拉全量 info
#    （本库全量约 89MB，服务端繁忙时需 10-30s）。
# 6. 服务端大量种子同时校验时偶发连接重置或查询长尾，查询需分片重试；
#    轮询保证至少 MIN_POLL_ROUNDS 轮完整查询，防止长尾吃掉超时窗口造成误判。

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
DELETE_WAIT_TIMEOUT = 60  # 等待删除真正完成的上限（整批共享，不是每个种子各等一次）
ADD_WAIT_TIMEOUT = 60     # 等待重加并进入做种态的上限（整批共享）
MIN_POLL_ROUNDS = 2       # 无论墙钟超时与否，至少完成几轮完整查询再下结论
QUERY_RETRIES = 3         # 单次查询失败后的重试次数（服务端繁忙时偶发连接重置）
QUERY_RETRY_PAUSE = 1.0   # 查询重试间隔
ADD_RETRY_LIMIT = 3       # 重加未能落地时最多尝试几轮
RETRY_PAUSE = 2           # 每轮重试前的共享退避，给 libtorrent 完成上一轮移除留出时间
DEFAULT_BATCH_SIZE = 50   # 每批种子数；设为 1 即退化为逐条处理
QUERY_SHARD_SIZE = 50     # 单次 info?hashes= 查询携带的 hash 上限


def is_ok(response):
    """成功判定用 2xx：5.3.0 的 torrents/add 在 pending 时返回 202，写死 200 会把成功当失败"""
    return 200 <= response.status_code < 300


def chunked(items, size):
    """按固定大小切片"""
    for start in range(0, len(items), size):
        yield items[start:start + size]


class QueryUnavailable(RuntimeError):
    """状态查询重试后仍失败：结果未知，不能据此判定种子失败"""


def query_states(client, hashes):
    """
    批量查询多个种子的状态，返回 {hash: torrent}。

    用 hashes 精确查询而非拉全量（本库 4 万多种子，全量响应约 89MB 且服务端繁忙时
    需 10-30s）。实测一次带 283 个 hash 的请求耗时显著高于分片，故按 QUERY_SHARD_SIZE
    切分；调用方每轮只传仍未确认的 hash，查询量随收敛递减。

    服务端在大量种子同时校验时偶发连接重置，故每个分片单独重试；全部耗尽时抛
    QueryUnavailable，由调用方按「结果未知」处理，避免把查询故障误判为种子失败。
    """
    states = {}
    for shard in chunked(list(hashes), QUERY_SHARD_SIZE):
        if not shard:
            continue
        last_error = None
        for attempt in range(QUERY_RETRIES):
            try:
                response = client.get('/api/v2/torrents/info', params={'hashes': '|'.join(shard)})
                if not is_ok(response):
                    raise RuntimeError(f'HTTP {response.status_code} {response.text[:200]}')
                for torrent in response.json():
                    states[torrent['hash']] = torrent
                last_error = None
                break
            except Exception as error:
                last_error = error
                if attempt < QUERY_RETRIES - 1:
                    time.sleep(QUERY_RETRY_PAUSE)
        if last_error is not None:
            raise QueryUnavailable(f'状态查询重试 {QUERY_RETRIES} 次仍失败: {last_error}')
    return states


def wait_batch_removed(client, hashes, timeout=DELETE_WAIT_TIMEOUT):
    """
    收敛式等待整批种子从会话中消失，返回 (仍残留的 hash 集合, 查询是否可靠)。

    libtorrent 移除是异步的，未消失就重加会撞 duplicate 被静默丢弃。
    整批共享一个等待窗口：每轮只查仍残留的 hash，而不是每个种子各等一轮。
    墙钟超时可能因单次查询长尾而提前触发，故保证至少完成 MIN_POLL_ROUNDS 轮。
    """
    pending = set(hashes)
    deadline = time.time() + timeout
    rounds = 0
    while pending:
        try:
            states = query_states(client, pending)
        except QueryUnavailable as error:
            print(f'⚠️ {error}')
            return pending, False
        pending = set(states) & pending
        rounds += 1
        if not pending or (rounds >= MIN_POLL_ROUNDS and time.time() >= deadline):
            break
        time.sleep(POLL_INTERVAL)
    return pending, True


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


def verify_batch(client, hashes, timeout=ADD_WAIT_TIMEOUT):
    """
    整批统一复核，返回 {hash: (结论, 说明)}。

    与逐条复核的区别：所有种子共用同一个等待窗口，每轮只对未定论的 hash 发一次
    批量查询；已定论的立即移出集合，后续轮次不再查它们。因此总耗时与种子数量
    解耦，只取决于最慢那个种子，而不是 N 个等待窗口相加。

    只认明确的完成态（DONE_STATES）：刚重加时服务端会短暂返回 checkingResumeData，
    此时进度已经是 1，若改用「非校验态 + 进度 100%」判定会把未跳过的种子误判为成功。
    """
    pending = set(hashes)
    verdicts = {}
    seen = set()
    last_state = {}
    deadline = time.time() + timeout
    rounds = 0
    query_reliable = True

    while pending:
        try:
            states = query_states(client, pending)
        except QueryUnavailable as error:
            # 查询故障不等于种子未落地，标为 suspect 交由人工/下一轮重试确认
            print(f'⚠️ {error}')
            query_reliable = False
            break
        for torrent_hash in list(pending):
            current = states.get(torrent_hash)
            if current is None:
                continue
            seen.add(torrent_hash)
            state = current.get('state', '')
            last_state[torrent_hash] = state
            if state in BAD_STATES:
                verdicts[torrent_hash] = ('failed', f'状态异常: {state}')
                pending.discard(torrent_hash)
            elif state in DONE_STATES:
                verdicts[torrent_hash] = ('ok', state)
                pending.discard(torrent_hash)
        rounds += 1
        if not pending or (rounds >= MIN_POLL_ROUNDS and time.time() >= deadline):
            break
        time.sleep(POLL_INTERVAL)

    for torrent_hash in pending:
        if not query_reliable:
            verdicts[torrent_hash] = ('unknown', '状态查询失败，结果未知')
        elif torrent_hash not in seen:
            verdicts[torrent_hash] = ('missing', '重加后未出现在会话中')
        else:
            verdicts[torrent_hash] = ('suspect', f'等待超时，仍在校验: {last_state.get(torrent_hash, "")}')
    return verdicts


def add_batch(client, records):
    """
    逐个提交 add（每个种子的回填参数不同，无法合并为单次请求），但提交后不等待。

    返回 {hash: 是否受理}，复核统一交给 verify_batch。
    """
    submitted = {}
    for record in records:
        accepted, detail = reimport_with_skip_check(client, record['torrent'], record['file'])
        record['last_detail'] = detail
        submitted[record['torrent']['hash']] = accepted
    return submitted


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


def skip_reimport(client, torrents, backup_dir, batch_size=DEFAULT_BATCH_SIZE):
    """
    分阶段流水线批量跳过校验（每批内「先全部添加完，再整批统一检查」）：

      1. 导出并预落盘整批 .torrent（含 tracker 与 passkey，删除后无原件则无法恢复）
      2. 一次请求批量删除整批
      3. 整批共享一个等待窗口，确认全部从会话消失
      4. 逐个提交 add，但提交后不等待
      5. 整批收敛式统一复核，已定论的立即移出后续查询集合
      6. 仅对未落地的少数条目做重试，退避同样整批共享

    返回 {'ok': [...], 'suspect': [...], 'failed': [...]}。未确认成功的条目保留
    .torrent 备份；已确认成功的备份自动清理，避免目录堆积。
    """
    results = {'ok': [], 'suspect': [], 'failed': []}
    total = len(torrents)
    processed = 0
    batch_count = (total + batch_size - 1) // batch_size

    for batch_no, batch in enumerate(chunked(torrents, batch_size), 1):
        print(f'--- 批次 {batch_no}/{batch_count}：{len(batch)} 个种子 ---')

        # 阶段 1：导出并预落盘。导出失败的不参与后续删除，避免无路可退
        records = []
        for torrent in batch:
            response = client.get('/api/v2/torrents/export', params={'hash': torrent['hash']})
            if not is_ok(response) or not response.content or not response.content.startswith(b'd'):
                print(f'❌ 获取种子文件失败，未做任何改动: {torrent["name"]} (HTTP {response.status_code})')
                results['failed'].append(torrent)
                continue
            torrent_file = response.content
            backup_path = save_torrent_backup(backup_dir, torrent['name'], torrent['hash'], torrent_file)
            records.append({'torrent': torrent, 'file': torrent_file, 'backup': backup_path})

        if not records:
            processed += len(batch)
            continue

        by_hash = {r['torrent']['hash']: r for r in records}

        # 阶段 2：一次请求批量删除
        hashes = list(by_hash)
        response = client.post('/api/v2/torrents/delete',
                               data={'hashes': '|'.join(hashes), 'deleteFiles': 'false'})
        if not is_ok(response):
            print(f'❌ 批量删除失败 (HTTP {response.status_code} {response.text[:200]})，本批未重加')
            for record in records:
                results['failed'].append(record['torrent'])
                print(f'    种子文件已备份: {record["backup"]}')
            processed += len(batch)
            continue

        # 阶段 3：整批共享一个等待窗口
        still_present, query_reliable = wait_batch_removed(client, hashes)
        if still_present:
            reason = (f'删除在 {DELETE_WAIT_TIMEOUT}s 内未完成' if query_reliable
                      else '删除状态查询失败，无法确认是否完成')
            for torrent_hash in still_present:
                record = by_hash.pop(torrent_hash)
                print(f'❌ {reason}，暂不重加: {record["torrent"]["name"]}')
                print(f'    种子文件已备份: {record["backup"]}')
                results['failed'].append(record['torrent'])

        if not by_hash:
            processed += len(batch)
            continue

        # 阶段 4-6：提交、整批复核、仅对未落地项重试
        pending = list(by_hash.values())
        final_verdicts = {}
        for attempt in range(1, ADD_RETRY_LIMIT + 1):
            submitted = add_batch(client, pending)
            # 未被受理且非 409 的失败无需重试；409 与 missing 都是删除竞态，可重试
            retryable = []
            for record in pending:
                torrent_hash = record['torrent']['hash']
                if not submitted[torrent_hash]:
                    if record['last_detail'].startswith('重加被拒绝（409'):
                        retryable.append(record)
                    else:
                        final_verdicts[torrent_hash] = ('failed', record['last_detail'])

            checkable = [r['torrent']['hash'] for r in pending if submitted[r['torrent']['hash']]]
            verdicts = verify_batch(client, checkable) if checkable else {}
            for torrent_hash, (verdict, detail) in verdicts.items():
                # missing = add 被静默丢弃；unknown = 查询故障，两者都值得整批重试
                if verdict in ('missing', 'unknown'):
                    retryable.append(by_hash[torrent_hash])
                else:
                    final_verdicts[torrent_hash] = (verdict, detail)

            pending = retryable
            if not pending:
                break
            if attempt < ADD_RETRY_LIMIT:
                print(f'♻️ 第 {attempt} 轮有 {len(pending)} 个种子未确认成功，{RETRY_PAUSE}s 后整批重试')
                time.sleep(RETRY_PAUSE)

        for record in pending:
            final_verdicts[record['torrent']['hash']] = (
                'suspect', f'{ADD_RETRY_LIMIT} 轮重加后仍无法在会话中确认: {record["last_detail"]}')

        # 汇总本批：成功清理备份，其余保留并留档
        for torrent_hash, record in by_hash.items():
            verdict, detail = final_verdicts.get(torrent_hash, ('suspect', '未获得复核结论'))
            torrent = record['torrent']
            if verdict == 'ok':
                results['ok'].append(torrent)
                try:
                    os.remove(record['backup'])
                except OSError:
                    pass
                continue
            if verdict == 'suspect':
                print(f'⚠️ 结果存疑，需人工确认: {torrent["name"]}（{detail}）')
                results['suspect'].append(torrent)
            else:
                print(f'❌ 重加未成功: {torrent["name"]}（{detail}）')
                results['failed'].append(torrent)
            print(f'    种子文件已备份，可手动重新导入: {record["backup"]}')

        processed += len(batch)
        print(f'--- 批次 {batch_no} 完成：累计 {processed}/{total} ---')

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
                   help='存放导出 .torrent 的目录（确认成功的备份会自动清理）')
    p.add_argument('--batch-size', type=int, default=DEFAULT_BATCH_SIZE,
                   help=f'每批处理的种子数，默认 {DEFAULT_BATCH_SIZE}；设为 1 退化为逐条处理')
    p.add_argument('--limit', type=int, default=0,
                   help='最多处理多少个种子，0 表示不限制；用于小样本试验')
    
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
    changed = False

    # 跳过校验中的种子
    if args.skip_check:
        checking_torrents = [t for t in torrents if t.get('state') in CHECKING_STATES]
        if args.limit > 0:
            checking_torrents = checking_torrents[:args.limit]
        print(f'待跳过校验: {len(checking_torrents)} 个'
              + (f'（--limit {args.limit} 小样本）' if args.limit > 0 else ''))
        if checking_torrents:
            report_skip_result('跳过校验', skip_reimport(client, checking_torrents,
                                                        args.backup_dir, args.batch_size))
            changed = True

    # 跳过强制做种的种子：forcedUP 与 checking 互斥，不受上一批影响，可复用同一份快照
    if args.skip_forced:
        forced_torrents = [t for t in torrents if t.get('state') == 'forcedUP']
        if args.limit > 0:
            forced_torrents = forced_torrents[:args.limit]
        print(f'待解除强制做种: {len(forced_torrents)} 个')
        if forced_torrents:
            report_skip_result('解除强制做种', skip_reimport(client, forced_torrents,
                                                            args.backup_dir, args.batch_size))
            changed = True

    # 仅在确实改动过种子后刷新一次全量快照；全量 info 约 89MB 且服务端繁忙时需 10-30s
    if changed:
        torrents = fetch_torrents(client)

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

    if args.limit > 0:
        rows = rows[:args.limit]

    df = pd.DataFrame(rows)
    df.to_excel(args.out, index=False)
    print(f'导出完成: {args.out} (共 {len(rows)} 条)')

if __name__ == '__main__':
    main()