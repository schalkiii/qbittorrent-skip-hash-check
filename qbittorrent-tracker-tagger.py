#!/usr/bin/env python3

# 说明：根据 tracker 返回信息对 qBittorrent 中的种子分类并打标签（仅打标签，绝不执行任何删除操作）
# 主要功能：
# 1. 遍历所有种子，读取每个种子的 tracker 状态与返回消息
# 2. 按预置规则将 tracker 错误/警告归类（已删种、审核未通过、站点不可达等 7 类）
# 3. 为命中的种子打上对应中文标签，并将明细导出到 Excel

# 鉴权：支持 API token（qBittorrent >= 5.2.0，推荐）与 用户名+密码，同时提供时 token 优先
#
# 用法示例（用 --token 或 --username/--password 均可）：
# 1. 预览模式（不打标签，仅导出 Excel）：
#    python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --dry-run
# 2. 默认模式（打标签 + 导出 Excel）：
#    python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx
# 3. 重新分类前先清理本脚本此前打的标签：
#    python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --clear-old
# 4. 一个种子命中多类时全部打上（默认只打优先级最高的一个）：
#    python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --multi-tag

import sys
import os

# 保证可从任意工作目录导入同目录下的 qb_auth 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse
import re
import time
import pandas as pd

from qb_auth import add_auth_args, build_client_from_args

# Windows 控制台默认 GBK 编码，无法输出 emoji 等字符，统一改为 UTF-8 并对不可编码字符降级
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# 标签统一前缀，便于 --clear-old 精确识别本脚本产生的标签
TAG_PREFIX = '错误/'

# 单次请求失败后的重试次数与间隔，应对 qb 大量种子时偶发的 502
REQUEST_RETRIES = 3
REQUEST_RETRY_DELAY = 1.0

# tracker 状态码含义（qBittorrent WebAPI v2）
# 实测 v5.3.0：0=禁用 1=未联系 2=工作中 3=更新中 4=不工作 5=错误
# 注意：部分版本用 4 表示 NotWorking，另一些（含 5.3.0）出现 5 表示错误，
# 因此错误判定统一用 >= 4，避免漏掉任何版本的错误状态。
TRACKER_STATUS_DISABLED = 0
TRACKER_STATUS_NOT_CONTACTED = 1
TRACKER_STATUS_WORKING = 2
TRACKER_STATUS_UPDATING = 3
TRACKER_STATUS_ERROR_MIN = 4  # 4 与 5 均视为错误

# 分类规则表，按列表顺序即为匹配优先级（越靠前优先级越高）
# 每项：(标签名, 该类含义说明, 关键词正则列表)
# 关键词统一用小写匹配，中文关键词直接子串匹配
CATEGORY_RULES = [
    (
        '已删种',
        '站点已删种或封禁该种子，本地数据已失去做种价值',
        [
            r'torrent not registered',
            r'torrent not exists',
            r'torrent does not exist',
            r'torrent banned',
            r'torrent deleted',
            r'unregistered torrent',
            r'没有在我们的\s*tracker\s*上注册',
            r'种子未注册',
            r'种子不存在',
            r'种子已删除',
        ],
    ),
    (
        '审核未通过',
        '种子处于候选/待审状态，站点侧尚未通过审核',
        [
            r'review not approved',
            r'not approved',
            r'审核',
        ],
    ),
    (
        '站点不可达',
        'DNS 解析失败、站点关站或网络不通',
        [
            r'不知道这样的主机',
            r'unreachable',
            r'name or service not known',
            r'no such host',
            r'stream truncated',
            r'timed out',
            r'timeout',
            r'connection refused',
            r'connection reset',
            r'connection aborted',
            r'temporary failure in name resolution',
            r'由于连接方在一段时间后没有正确答复或连接的主机没有反应',
            r'不知道这样的主机。',
            r'由于连接尝试失败',
            r'由本地系统中止网络连接',
            r'目标计算机积极拒绝',
            r'通常每个套接字地址',
        ],
    ),
    (
        '无权限',
        'passkey 失效、账号被封、客户端不被站点允许或缺少鉴权参数',
        [
            r'forbidden',
            r'unauthorized',
            r'require passkey',
            r'invalid passkey',
            r'passkey',
            r'access denied',
            r'not allow',
            r'not allowed',
            r'client.*not.*allow',
            r'is banned',
            r'banned client',
        ],
    ),
    (
        '服务端异常',
        '站点服务器临时故障或返回了无法解析的响应，通常可自行恢复',
        [
            r'bad gateway',
            r'service unavailable',
            r'internal server error',
            r'gateway time-?out',
            r'not found',
            r'<none>',
            r'bencoded string',
            r'invalid response',
            r'malformed',
            r'502',
            r'503',
        ],
    ),
    (
        '汇报过频',
        '汇报间隔过短或多个同站 tracker 共用 passkey，一般无害',
        [
            r'too frequent',
            r'minimum announce time',
            r'too many requests',
            r'duplicate report',
            r'汇报过了',
            r'req too frequently',
            r'announce time error',
        ],
    ),
    (
        '其他',
        '未能归入上述类别的 tracker 消息，多数实为正常应答',
        [],  # 空列表表示兜底类，由代码显式处理
    ),
]

# 兜底类标签名，单独提出便于代码引用
FALLBACK_CATEGORY = CATEGORY_RULES[-1][0]

# 预编译正则，避免在数万种子的循环中重复编译
COMPILED_RULES = [
    (name, [re.compile(p, re.IGNORECASE) for p in patterns])
    for name, desc, patterns in CATEGORY_RULES
]

# 本脚本管理的全部标签名，用于 --clear-old
ALL_TAGS = [TAG_PREFIX + name for name, _, _ in CATEGORY_RULES]


def fetch_torrents(client):
    """拉取全部种子信息。41015 个种子时响应可达数十 MB，读取超时已放宽至 60s。"""
    r = client.get('/api/v2/torrents/info')
    r.raise_for_status()
    return r.json()


def fetch_trackers(client, torrent_hash):
    """
    读取单个种子的 tracker 列表。
    种子数量巨大时 qb 偶发 502，故加入有限重试而非直接放弃该种子。
    """
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


def classify_message(msg):
    """
    将单条 tracker 消息归类，返回标签名（不含前缀）。
    按 CATEGORY_RULES 顺序匹配，先命中者优先；均不命中则归入兜底类。
    """
    for name, patterns in COMPILED_RULES:
        for pattern in patterns:
            if pattern.search(msg):
                return name
    return FALLBACK_CATEGORY


def collect_issues(client, torrents, progress_step):
    """
    遍历种子读取 tracker 信息，收集所有状态异常的条目。
    异常定义：status=4（错误），或 msg 非空（警告）。
    返回 (明细行列表, {hash: [标签名, ...]})，标签名按优先级升序排列。
    """
    rows = []
    tag_map = {}
    failed = []
    total = len(torrents)

    for index, t in enumerate(torrents):
        torrent_hash = t['hash']
        try:
            trackers = fetch_trackers(client, torrent_hash)
        except Exception as e:
            failed.append((t.get('name', ''), torrent_hash, str(e)))
            continue

        hit_names = []
        for tr in trackers:
            url = tr.get('url', '')
            # DHT / PeX / LSD 为伪 tracker 条目，跳过
            if url.startswith('**'):
                continue
            status = tr.get('status')
            msg = (tr.get('msg') or '').strip()
            # 错误：状态码 4 或 5（部分版本用 5 表示错误）
            # 警告：状态码非错误但 tracker 返回了消息（如汇报过频）
            is_error = (status >= TRACKER_STATUS_ERROR_MIN)
            if not is_error and not msg:
                continue

            # 空消息但状态为错误：视为服务端异常（tracker 未给出具体原因）
            if is_error and not msg:
                name = '服务端异常'
            else:
                name = classify_message(msg)
            if name not in hit_names:
                hit_names.append(name)

            rows.append({
                'name': t.get('name', ''),
                'hash': torrent_hash,
                'torrent_state': t.get('state', ''),
                'tracker_url': url,
                'tracker_status': status,
                'tracker_level': '错误' if is_error else '警告',
                'tracker_msg': msg,
                'category': name,
                'tag': TAG_PREFIX + name,
                'save_path': t.get('save_path', t.get('content_path', '')),
                'size_readable': nice_bytes(t.get('size', t.get('total_size', 0))),
                'ratio': t.get('ratio', ''),
                'existing_tags': t.get('tags', ''),
                'added_on': (time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t.get('added_on', 0)))
                             if t.get('added_on') else ''),
            })

        if hit_names:
            # 按 CATEGORY_RULES 中的顺序排序，保证优先级语义稳定
            order = {name: i for i, (name, _, _) in enumerate(CATEGORY_RULES)}
            tag_map[torrent_hash] = sorted(hit_names, key=lambda n: order[n])

        if progress_step and (index + 1) % progress_step == 0:
            print(f'  已扫描 {index + 1}/{total}', flush=True)

    return rows, tag_map, failed


def create_tags(client, tags):
    """预创建标签，避免部分 qb 版本在 addTags 时静默失败"""
    r = client.post('/api/v2/torrents/createTags', data={'tags': ','.join(tags)})
    return r.status_code == 200


def add_tags(client, hashes, tag):
    """为一批种子添加同一个标签"""
    r = client.post('/api/v2/torrents/addTags', data={'hashes': '|'.join(hashes), 'tags': tag})
    return r.status_code == 200


def remove_tags(client, hashes, tag):
    """从一批种子上移除同一个标签（仅解除关联，不影响种子本身）"""
    r = client.post('/api/v2/torrents/removeTags', data={'hashes': '|'.join(hashes), 'tags': tag})
    return r.status_code == 200


def clear_old_tags(client, torrents, batch_size):
    """
    清理本脚本此前打的标签。
    仅对当前带有这些标签的种子执行 removeTags，不做任何删除操作。
    """
    cleared = 0
    for tag in ALL_TAGS:
        hashes = [t['hash'] for t in torrents
                  if tag in {x.strip() for x in (t.get('tags') or '').split(',') if x.strip()}]
        if not hashes:
            continue
        for i in range(0, len(hashes), batch_size):
            batch = hashes[i:i + batch_size]
            if remove_tags(client, batch, tag):
                cleared += len(batch)
            else:
                print(f'[失败] 移除标签: {tag} (共 {len(batch)} 个)')
    return cleared


def apply_tags(client, tag_map, multi_tag, batch_size):
    """
    按标签分组批量打标签。
    multi_tag=False 时每个种子只打优先级最高的标签。
    """
    # 按标签聚合 hash，减少 API 调用次数
    grouped = {}
    for torrent_hash, names in tag_map.items():
        selected = names if multi_tag else names[:1]
        for name in selected:
            grouped.setdefault(TAG_PREFIX + name, []).append(torrent_hash)

    if not grouped:
        return {}

    create_tags(client, list(grouped.keys()))

    result = {}
    for tag, hashes in grouped.items():
        tagged = 0
        for i in range(0, len(hashes), batch_size):
            batch = hashes[i:i + batch_size]
            if add_tags(client, batch, tag):
                tagged += len(batch)
            else:
                print(f'[失败] 打标签: {tag} (共 {len(batch)} 个)')
        result[tag] = tagged
        print(f'[完成] {tag}: {tagged} 个种子')
    return result


def nice_bytes(n):
    try:
        n = int(n)
    except:
        return str(n)
    for unit in ['B', 'KB', 'MB', 'GB', 'TB', 'PB']:
        if n < 1024:
            return f"{n:.2f}{unit}"
        n /= 1024
    return f"{n:.2f}EB"


def export_excel(rows, tag_map, multi_tag, out_path):
    """
    导出两个工作表：
    - 明细：每条异常 tracker 一行
    - 汇总：按分类统计条目数与种子数
    """
    detail_df = pd.DataFrame(rows)

    order = {name: i for i, (name, _, _) in enumerate(CATEGORY_RULES)}
    desc_map = {name: desc for name, desc, _ in CATEGORY_RULES}

    # 统计每个分类实际会被打标签的种子数（受 multi_tag 影响）
    tagged_count = {}
    for names in tag_map.values():
        for name in (names if multi_tag else names[:1]):
            tagged_count[name] = tagged_count.get(name, 0) + 1

    summary_rows = []
    for name, desc, _ in CATEGORY_RULES:
        sub = [r for r in rows if r['category'] == name]
        if not sub and not tagged_count.get(name):
            continue
        summary_rows.append({
            'category': name,
            'tag': TAG_PREFIX + name,
            'description': desc,
            'tracker_entries': len(sub),
            'error_entries': len([r for r in sub if r['tracker_level'] == '错误']),
            'warn_entries': len([r for r in sub if r['tracker_level'] == '警告']),
            'involved_torrents': len({r['hash'] for r in sub}),
            'tagged_torrents': tagged_count.get(name, 0),
        })
    summary_rows.sort(key=lambda r: order[r['category']])
    summary_df = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='汇总', index=False)
        detail_df.to_excel(writer, sheet_name='明细', index=False)

    return summary_df


def main():
    p = argparse.ArgumentParser(
        description='根据 tracker 返回信息对 qBittorrent 种子分类打标签（只打标签，不做任何删除）')
    p.add_argument('--host', '-H', required=True, help='qBittorrent Web UI 地址，例如 http://127.0.0.1:8080')
    add_auth_args(p)
    p.add_argument('--out', '-o', default='qb_tracker_issues.xlsx', help='输出Excel文件名')
    p.add_argument('--dry-run', action='store_true', help='预览模式：只导出Excel，不打任何标签')
    p.add_argument('--clear-old', action='store_true', help='打标签前先清理本脚本此前打的标签')
    p.add_argument('--multi-tag', action='store_true', help='一个种子命中多类时全部打上（默认只打优先级最高的一个）')
    p.add_argument('--batch-size', type=int, default=200, help='批量打标签时每批种子数量')
    p.add_argument('--progress-step', type=int, default=5000, help='扫描进度打印间隔，0 表示不打印')

    args = p.parse_args()
    client = build_client_from_args(args)
    print(f'鉴权方式: {"API token" if client.auth_mode == "token" else "用户名密码(Cookie 会话)"}')

    try:
        run(client, args)
    finally:
        # Cookie 模式下主动登出，避免服务端会话堆积；token 模式无状态，无需处理
        client.logout()

def run(client, args):
    torrents = fetch_torrents(client)
    print(f'共 {len(torrents)} 个种子，开始扫描 tracker 状态...')

    if args.clear_old:
        if args.dry_run:
            print('预览模式：跳过清理旧标签')
        else:
            cleared = clear_old_tags(client, torrents, args.batch_size)
            print(f'已清理旧标签关联 {cleared} 次')
            torrents = fetch_torrents(client)  # 重新获取最新标签状态

    rows, tag_map, failed = collect_issues(client, torrents, args.progress_step)

    if failed:
        print(f'\n[警告] {len(failed)} 个种子的 tracker 读取失败（重试后仍失败），已跳过：')
        for name, torrent_hash, err in failed[:10]:
            print(f'  - {name[:50]} ({torrent_hash[:12]}): {err}')
        if len(failed) > 10:
            print(f'  ... 其余 {len(failed) - 10} 个省略')

    if not rows:
        print('未发现任何 tracker 错误或警告。')
        return

    error_entries = len([r for r in rows if r['tracker_level'] == '错误'])
    warn_entries = len(rows) - error_entries
    print(f'\n扫描完成：错误 {error_entries} 条，警告 {warn_entries} 条，涉及 {len(tag_map)} 个种子')

    summary_df = export_excel(rows, tag_map, args.multi_tag, args.out)
    print(f'\n分类汇总：')
    print(summary_df.to_string(index=False))
    print(f'\n导出完成: {args.out}')

    if args.dry_run:
        print('\n预览模式：未对 qBittorrent 做任何修改。确认无误后去掉 --dry-run 即可打标签。')
        return

    print('\n开始打标签...')
    apply_tags(client, tag_map, args.multi_tag, args.batch_size)
    print('打标签完成（未执行任何删除操作）')


if __name__ == '__main__':
    main()
