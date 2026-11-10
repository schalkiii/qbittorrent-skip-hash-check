# AGENTS.md

## 项目定位

qBittorrent WebAPI 批量运维脚本集：跳过种子校验、按 tracker 错误分类打标签、批量替换 tracker 域名。

## 怎么跑起来

```bash
pip install requests pandas openpyxl
# 支持 API token（>=5.2.0，推荐）或 用户名+密码，token 优先
python qbittorrent-skip-hash-check.py --host http://192.168.10.111:9091 --token qbt_xxx --filter checking --out checking.xlsx
```

`.bat` 一键脚本含连接凭据，已在 `.gitignore` 中排除，不入库。

## 技术栈

Python 3 + requests + pandas + openpyxl。无测试、无构建，直接运行脚本。

## 目录与约定

| 文件 | 作用 |
|---|---|
| `qb_auth.py` | 共享鉴权模块，`QBClient` 封装全部 HTTP 调用 |
| `qbittorrent-skip-hash-check.py` | 跳过校验 / 解除强制做种（删除后重导入） |
| `qbittorrent-tracker-tagger.py` | 按 tracker 错误信息分类打标签（7 类，只打标签不删除） |
| `qbittorrent-tracker-url-replacer.py` | 批量替换 tracker 域名，保留 passkey |

## 必守的坑（下次别再踩）

- **成功判定用 2xx，不要写死 200**：该版本 `auth/login` 与 `torrents/editTracker` 成功均返回 **204** 且响应体为空。
- **tracker 错误状态码不要写死 4**：实测 v5.3.0 中错误状态可能是 **5**，错误判定统一用 `status >= 4`。
- **token 鉴权必须 `Authorization: Bearer`**：实测 `X-API-Key` 头在 v5.3.0 返回 403。
- **token 不能用于 `auth/` 端点**（403），验证 token 改为探测 `/api/v2/app/version`。
- **`editTracker` 参数是 `hash` + `url`（旧地址）+ `newUrl`**：`origUrl` 属于 editWebSeed，别混用。
- **跳过校验的参数名在 5.3.0 变了**：`torrents/add` 的 `skip_checking` 改名 `seedMode`，且服务端对不认识的参数**静默忽略**（HTTP 依然成功）。新旧名字必须一起发。
- **`paused` 在 5.3.0 改名 `stopped`**：不发时跟随全局偏好 `add_stopped_enabled`（本机为 `true`），种子会被直接加成 `stoppedUP`，即用户看到的「暂停状态」。同样双发，并显式置 `false`。
- **`torrents/delete` 是异步的**：`removeTorrent` 只把种子摘出会话，libtorrent 移除完成前重加同 hash 会被 `duplicate_is_error` 静默丢弃（上游 #24367）。必须轮询 `torrents/info?hash=` 确认消失后再重加。
- **`torrents/add` 现在返回 JSON 计数**：`success_count`/`failure_count`/`pending_count`/`added_torrent_ids`；有 pending 时 HTTP **202**，一条都没加时 **409**。只判 `== 200` 会把成功当失败。
- **AutoTMM 打开时不要发 `savepath`**：服务端见 `savepath` 会隐式关闭 AutoTMM（`initLoadTorrentParams`），改发 `autoTMM=true`。
- Windows 控制台是 GBK，脚本打印 emoji 前必须 `sys.stdout.reconfigure(encoding='utf-8', errors='replace')`。
- 大量种子（4 万+）时全量 `torrents/info` 约 89MB，服务端繁忙需 10-30s；复核必须用 `torrents/info?hashes=` 按 50 个分片，勿拉全量。
- **不要逐条轮询等待**：跳过校验采用批量流水线（先全部 add、再整批收敛复核），每轮只查未定论的 hash；逐条复核会让 N 个 0.5s 等待窗口相加。
- 服务端大量种子同时校验时偶发连接重置与查询长尾（0.04-26.7s），查询需分片重试，并保证至少 2 轮完整查询再下结论。

## 当前状态

- qBittorrent v5.3.0 / WebAPI 2.16.2，约 4.1 万种子。
- 2026-11-01：三个脚本统一接入 `qb_auth.py` 鉴权；新增 tracker 域名批量替换脚本并完成 `t.hddolby.com` → `www.uhddolby.com` 共 408 个种子迁移。
- 2026-11-01：修复 tagger 漏标问题——错误状态码判定改为 `>= 4`，`已删种` 增补中文模式（`种子不存在` 等），
  `站点不可达` 增补中文连接错误，`无权限` 增补客户端封禁，`服务端异常` 增补 `Not Found`/bencode 失败。
- 2026-11-07：修复 5.3.0 升级后跳过校验失效（`seedMode`/`stopped` 改名 + 异步删除竞态 + 202/409 判定）；
  重加时回填分类/标签/限速/分享率等属性，并增加复核与失败留档 `skip_check_backup/`。
  实测 8.09GB 种子 3.6s 完成且属性无损；**注意本地做种时长与上传量会被删除重加清零**。
- 2026-11-10：跳过校验改为批量流水线（`--batch-size` 默认 50、`--limit` 小样本），
  先全部 add 再整批收敛复核；消除逐条串行等待与重复全量 `info` 拉取。
