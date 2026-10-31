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
- Windows 控制台是 GBK，脚本打印 emoji 前必须 `sys.stdout.reconfigure(encoding='utf-8', errors='replace')`。
- 大量种子（4 万+）时 `torrents/info` 响应数十 MB，读取超时需放宽到 60s。

## 当前状态

- qBittorrent v5.3.0 / WebAPI 2.16.2，约 4.1 万种子。
- 2026-11-01：三个脚本统一接入 `qb_auth.py` 鉴权；新增 tracker 域名批量替换脚本并完成 `t.hddolby.com` → `www.uhddolby.com` 共 408 个种子迁移。
- 2026-11-01：修复 tagger 漏标问题——错误状态码判定改为 `>= 4`，`已删种` 增补中文模式（`种子不存在` 等），
  `站点不可达` 增补中文连接错误，`无权限` 增补客户端封禁，`服务端异常` 增补 `Not Found`/bencode 失败。
