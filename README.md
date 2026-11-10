# qBittorrent 工具集

本仓库包含三个 qBittorrent Web API 辅助脚本：跳过校验、tracker 错误分类打标签、tracker 域名批量替换。

## 脚本一览

| 脚本                            | 用途                                                     |
| ------------------------------- | -------------------------------------------------------- |
| `qbittorrent-skip-hash-check.py` | 跳过种子校验 / 解除强制做种，支持预览与执行模式           |
| `qbittorrent-tracker-tagger.py`  | 根据 tracker 返回的错误/警告信息分类种子并打标签          |

---

# qBittorrent Skip Hash Check

跳过 qBittorrent 校验的工具，支持预览模式和实际执行模式。

## 功能特性

- **预览模式**：导出指定状态的种子信息到 Excel，便于确认
- **跳过校验**：将校验中的种子改为做种状态（通过删除后重新导入实现）
- **解除强制做种**：将强制做种的种子改为普通做种状态
- **批量处理**：支持一次性处理多个种子

### 组件

| 文件 | 说明 |
| ---- | ---- |
| `qb_auth.py` | 共享鉴权模块，三个脚本复用 |
| `qbittorrent-skip-hash-check.py/.bat` | 跳过校验 / 解除强制做种 |
| `qbittorrent-tracker-tagger.py/.bat` | 按 tracker 错误信息分类打标签（只打标签，不删除） |
| `qbittorrent-tracker-url-replacer.py/.bat` | 批量替换 tracker 域名，保留 passkey |

`.bat` 一键脚本内含连接凭据，已在 `.gitignore` 中排除，不入库。

## 环境要求

- Python 3.6+
- 依赖库：`requests`、`pandas`、`openpyxl`

```bash
pip install requests pandas openpyxl
```

## 鉴权方式

支持两种方式，**同时提供时 API token 优先**（与服务端行为一致）。

### 1. API token（推荐，qBittorrent >= 5.2.0 / WebAPI >= v2.14.1）

在 WebUI → 设置 → Web UI → API Key 处生成，格式为 `qbt_` 前缀 + 28 位字符（共 32 字符）。

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter checking --out checking.xlsx
```

注意事项：

- 请求头为 `Authorization: Bearer <token>`。实测 `X-API-Key` 头在 v5.3.0 上返回 403，故统一使用 Bearer。
- token **不能用于 `auth/` 端点**（`login`/`logout` 会返回 403），因此脚本改用 `app/version` 探测 token 有效性。
- 重新生成 API key 会立即作废旧 key。
- 低版本 qBittorrent（< 5.2.0）不支持，会自动回退到用户名密码。

### 2. 用户名 + 密码（传统 Cookie 会话）

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password YOUR_PASSWORD
```

兼容新旧登录响应：旧版返回 `200 + "Ok."`，4.6+ 返回 `204` 且响应体为空。脚本结束时主动登出，避免服务端会话堆积。

### 关于性能

token 鉴权**对脚本单次运行的性能提升非常有限**，主要价值在可靠性：

| 维度 | 用户名密码 | API token |
| ---- | ---------- | --------- |
| 登录往返 | 1 次额外 HTTP 往返 | 无 |
| 服务端密码校验 | PBKDF2 哈希，约 5–50ms CPU | 常量时间字符串比较，微秒级 |
| 会话管理 | 每次运行新建会话并下发 Cookie | 以 token 为 ID 复用同一会话 |
| 失败封禁 | 连续失败达阈值会封禁来源 IP | 不触发失败计数 |
| 单次脚本运行总收益 | — | 约省 10–100ms（局域网） |

结论：对一次性脚本，收益在总耗时中占比可忽略（本库 4 万多种子，扫描耗时以分钟计）。真正受益的是**高频轮询 / 常驻集成**场景（每分钟请求一次可省下持续的登录开销与哈希 CPU），以及避免会话堆积和误封 IP。

## 使用方法

### 基本参数

| 参数            | 说明                                                |
| --------------- | --------------------------------------------------- |
| `--host`        | qBittorrent Web UI 地址，如 `http://127.0.0.1:8080` |
| `--token`       | API token（>= 5.2.0），优先于用户名密码             |
| `--username`    | 用户名（未提供 token 时必填）                       |
| `--password`    | 密码（未提供 token 时必填）                         |
| `--filter`      | 过滤类型：`checking`、`completed`、`forced`、`all`  |
| `--out`         | 输出 Excel 文件名                                   |
| `--skip-check`  | 跳过校验中的种子                                    |
| `--skip-forced` | 跳过强制做种的种子                                  |
| `--backup-dir`  | 存放 .torrent 的目录，默认 `skip_check_backup/`；确认成功的备份自动清理 |
| `--batch-size`  | 每批处理的种子数，默认 50；设为 1 退化为逐条处理 |
| `--limit`       | 最多处理多少个种子，0 表示不限制；用于小样本试验 |

### 过滤选项说明

| 选项        | 说明                 |
| ----------- | -------------------- |
| `checking`  | 校验中的种子（默认） |
| `completed` | 已完成的种子         |
| `forced`    | 强制做种的种子       |
| `all`       | 全部种子             |

### 使用示例

下面示例统一使用 `--token`，也可替换为 `--username admin --password 123456`。

**1. 预览模式 - 查看校验中的种子**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter checking --out checking.xlsx
```

**2. 执行模式 - 跳过校验**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter checking --out checking.xlsx --skip-check
```

**3. 预览模式 - 查看强制做种的种子**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter forced --out forced.xlsx
```

**4. 执行模式 - 解除强制做种**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --filter forced --out forced.xlsx --skip-forced
```

### Tracker Tagger 用法

```bash
REM 预览：只导出 Excel，不打标签
python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --dry-run

REM 打标签前先清理旧标签，再重新分类
python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --token qbt_xxxxxxxx --clear-old
```

| 参数           | 说明                                          |
| -------------- | --------------------------------------------- |
| `--dry-run`    | 预览模式，只导出 Excel，不打任何标签          |
| `--clear-old`  | 打标签前先清理本脚本此前打的标签              |
| `--multi-tag`  | 命中多类时全部打上（默认只打优先级最高的一个）|

#### 分类规则（7 类，按优先级排序）

| 分类 | 典型消息 |
| ---- | -------- |
| `已删种` | `Torrent not exists`、`PTT:种子不存在`、`torrent not registered`、`种子未注册` |
| `审核未通过` | `review not approved`、含"审核" |
| `站点不可达` | `不知道这样的主机`、`connection refused`、中文连接超时/拒绝 |
| `无权限` | `Forbidden`、`invalid passkey`、`Client ... is not allow` |
| `服务端异常` | `Bad Gateway`、`Not Found`、bencode 解析失败 |
| `汇报过频` | `Request too frequent`、`minimum announce time` |
| `其他` | 未匹配的兜底，多数实为正常应答 |

错误状态判定使用 `status >= 4`（部分 qB 版本用 5 表示错误，写死 4 会漏判）。

### Tracker URL Replacer 用法（站点换域名时迁移）

只替换 URL 的 **host 字段**，`passkey`、路径、查询串原样保留。

```bash
REM 1. 预览：列出将被修改的种子，不做任何改动
python qbittorrent-tracker-url-replacer.py --host http://127.0.0.1:8080 --token qbt_xxx \
    --old-host t.hddolby.com --new-host www.uhddolby.com

REM 2. 小批量试验（默认 2 个），改完立即复核
python qbittorrent-tracker-url-replacer.py --host ... --token qbt_xxx \
    --old-host t.hddolby.com --new-host www.uhddolby.com --canary

REM 3. 确认无误后批量执行
python qbittorrent-tracker-url-replacer.py --host ... --token qbt_xxx \
    --old-host t.hddolby.com --new-host www.uhddolby.com --apply
```

| 参数              | 说明                                        |
| ----------------- | ------------------------------------------- |
| `--old-host`      | 旧域名，如 `t.hddolby.com`                  |
| `--new-host`      | 新域名，如 `www.uhddolby.com`               |
| `--canary`        | 试验模式，只改前 N 个并复核（默认 2 个）    |
| `--apply`         | 批量执行（默认需交互输入 yes，`--yes` 跳过）|
| `--backup`        | 变更前计划备份 JSON，便于审计               |

安全设计：

- 只处理 tracker URL 命中旧域名的种子，其余一律不触碰
- 变更前导出完整映射备份 JSON
- 服务端保证不重复添加：新 URL 已存在时返回 Conflict，脚本记为「跳过」
- `editTracker` 只触发 `forceReannounce`，**不会 recheck/rehash**，已下载数据不受影响
- 执行后自动复核：确认新 URL 生效且旧 URL 已消失

注意：`editTracker` 的参数是 `hash` + **`url`**（旧地址）+ `newUrl`；`origUrl` 属于 `editWebSeed`，不可混用。

## 工作原理

该脚本通过以下步骤跳过种子校验：

1. **导出种子文件** — 从 qBittorrent 导出原始 .torrent（含 tracker 与 passkey）。导出失败或内容不是合法 bencode 时直接跳过该种子，不做任何改动
2. **删除种子** — 删除原种子记录（`deleteFiles=false`，保留文件）
3. **等待删除完成** — 5.x 的删除是异步的，轮询确认该 hash 已从会话消失后才继续；没消失就重加会被 libtorrent 判为 duplicate 而**静默丢弃**（HTTP 仍返回成功）
4. **重新导入** — 带上跳过校验参数重加，并显式要求立即启动
5. **逐条复核** — 轮询确认种子回到会话且进入做种态，才计入「成功」；仍在校验的记为「存疑」

### qBittorrent 5.3.0（WebAPI v2.16）的参数改名

| 用途 | ≤ 5.2 参数名 | 5.3.0 参数名 | 说明 |
| ---- | ------------ | ------------ | ---- |
| 跳过校验 | `skip_checking` | `seedMode` | 5.3.0 起只解析 `seedMode`，旧名字被静默忽略 → 表现为「重加成功但仍在校验」 |
| 加入后是否停止 | `paused` | `stopped` | 不发时跟随全局偏好 `add_stopped_enabled`；该偏好为 `true` 时新种子被加成 `stoppedUP`（即「暂停状态」） |

脚本对这两组**新旧名字一起发**，任何一端只会解析自己认识的那个，因此同一份脚本在 4.6 / 5.0 / 5.3.0 上都能用。

`torrents/add` 的返回也变了：成功返回 JSON 计数（`success_count` / `failure_count` / `pending_count` / `added_torrent_ids`），有 pending 时为 **202**，一条都没加进去时为 **409**。所以判定一律用 2xx 并解析 JSON，不能写死 200。

### 重加时回填的属性

分类、标签、保存路径、上传/下载限速、顺序下载、首尾优先、分享率限制、AutoTMM、显示名。
AutoTMM 打开时**不发** `savepath`——服务端看到 `savepath` 会隐式关掉 AutoTMM，路径改由分类/默认规则推导。

> `forced`（强制做种）故意不回填：解除强制做种正是 `--skip-forced` 的目的。

### 批量流水线（先全部添加，再统一检查）

跳过校验采用分阶段流水线，每批内**先把所有种子重新添加完，再整批统一复核**，而不是每加一个就轮询等它：

1. 导出并预落盘整批 `.torrent`（含 tracker 与 passkey）
2. 一次请求批量删除整批（`hashes` 用 `|` 分隔）
3. 整批共享一个等待窗口，确认全部从会话消失
4. 逐个提交 `add`（每个种子的回填参数不同，无法合并），提交后不等待
5. 整批收敛式统一复核：每轮只查未定论的 hash，已定论的立即移出
6. 仅对未落地的少数条目重试，退避整批共享

状态查询统一用 `torrents/info?hashes=` 精确查询并按 50 个分片：**全量 `info` 在本库约 89MB，服务端繁忙时需 10–30s**，不可用。实测一次带 283 个 hash 的请求耗时显著高于分片。

| 维度（N=283, batch=50） | 逐条复核（旧） | 批量流水线（新） |
| --- | --- | --- |
| `delete` 请求数 | 283 | 6 |
| 等待/复核查询 | ≥566 次单 hash 查询 | 约 12 次批量查询 |
| 串行 `sleep` | ≈283s | ≈6s |
| 全量 `info` 拉取 | 3 次 | 1 次 |

**可靠性加固**：服务端在大量种子同时校验时偶发连接重置或单次查询长尾（实测 0.04–26.7s），因此每个分片单独重试 3 次；查询故障标记 `unknown` 并整批重试，不会把查询故障误判为种子失败；轮询保证至少完成 2 轮完整查询才下结论，避免长尾吃掉超时窗口。

### 代价与兜底

- **本地统计会被清零**：删除重加后该种子的「已做种时间 / 本地上传量」从头计数（实测 8.09GB 种子重加后 `seeding_time` 只剩 7 秒）。站点侧统计走 announce，不受影响。
- **重加失败时自动留档**：导出失败之外的任何失败（删除卡住、409、未被受理）都会保留 `.torrent` 到 `--backup-dir`（默认 `skip_check_backup/`），可手动导入找回，passkey 不会丢；确认成功的备份自动清理。
- **小样本试验**：首次使用建议 `--limit 2` 先验证，确认无误后再全量执行。


## 导出字段说明

导出的 Excel 包含以下字段：

| 字段               | 说明                 |
| ------------------ | -------------------- |
| `name`             | 种子名称             |
| `hash`             | 种子哈希             |
| `state`            | 当前状态             |
| `progress`         | 完成进度 (%)         |
| `size_bytes`       | 文件大小（字节）     |
| `size_readable`    | 文件大小（可读格式） |
| `save_path`        | 保存路径             |
| `num_files`        | 文件数量             |
| `completed_pieces` | 已完成块数           |
| `ratio`            | 分享率               |
| `seeds`            | 做种数               |
| `peers`            | 连接数               |
| `added_on`         | 添加时间             |
| `last_activity`    | 最后活动时间         |
| `category`         | 分类                 |
| `trackers`         | trackers             |
| `label`            | 标签                 |
| `files`            | 文件列表             |

## 注意事项

1. **建议先使用预览模式**：确认无误后再添加 `--skip-check` 或 `--skip-forced` 执行
2. **只处理目标状态**：仅触碰 `checkingDL` / `checkingUP`（以及 `--skip-forced` 时的 `forcedUP`），其余种子一律不动
3. **保留文件**：删除种子时默认保留文件，不会删除实际数据
4. **本地统计会清零**：跳过校验靠删除重加实现，该种子的本地做种时长与上传量会重新开始计数；站点侧记录不受影响
5. **留意「存疑」计数**：汇总里的存疑表示重加后仍在慢慢校验，即跳过并未真正生效，需要人工确认而不是重复跑
6. **API 兼容性**：需要 qBittorrent 开启 Web API

---

# qBittorrent Tracker Tagger

根据 tracker 返回的错误/警告信息，将种子自动分类并打上中文层级标签（前缀 `错误/`），便于在 qBittorrent 中过滤排查。**脚本仅做只读查询与打标签（addTags/createTags），绝不执行任何删除操作。**

## 功能特性

- **只读安全**：仅调用 `torrents/info` 与 `torrents/trackers` 查询，以及 `addTags`/`createTags` 打标签，不删除任何种子或数据
- **错误与警告双维度**：依据 tracker 的 `status`（4=错误）与 `msg` 非空（警告）两类状态筛选
- **七类自动分类**：按优先级匹配报错文案，命中即打对应标签
- **导出 Excel**：将命中的 tracker 报错明细导出为 `qb_tracker_issues.xlsx`
- **可清理重跑**：`--clear-old` 先移除本脚本此前打的标签，再重新分类，避免标签堆积

## 分类规则

| 优先级 | 标签             | 匹配关键词（不区分大小写）                                               | 说明                       |
| ------ | ---------------- | ------------------------------------------------------------------------ | -------------------------- |
| 1      | `错误/已删种`    | `torrent not registered`、`torrent not exists`、`torrent does not exist`、 `torrent banned`、`torrent deleted`、中文「未注册」 | 站点已删种/封禁，本地数据失效 |
| 2      | `错误/审核未通过`| `torrent review not approved`                                            | 候选区待审，站点侧未过审   |
| 3      | `错误/站点不可达`| `不知道这样的主机`、`unreachable`、`Name or service not known`、`stream truncated` | 域名失效/关站/网络不通   |
| 4      | `错误/无权限`    | `Forbidden`、`Require passkey`、`Unauthorized`、`passkey`                 | passkey 失效或账号被禁     |
| 5      | `错误/服务端异常`| `Bad Gateway`、`Service Unavailable`、`Internal Server Error`、`Gateway Timeout` | 站点临时故障，可自愈     |
| 6      | `错误/汇报过频`  | `too frequent`、`minimum announce time`、`Too Many Requests`、`Duplicate report`、中文「汇报过了」 | 多 tracker 同 passkey 导致，可忽略 |
| 7      | `错误/其他`      | 兜底层，如 `ook`、`Working on another tracker`、`<none>`                  | 多为正常应答的误判，仅作记录 |

> 一个种子若命中多类，默认只打优先级最高的一个标签（可用 `--multi-tag` 全部打上）。`错误/已删种` 优先级最高。

## 参数说明

| 参数               | 说明                                                                         |
| ------------------ | ---------------------------------------------------------------------------- |
| `--host`           | qBittorrent Web UI 地址，如 `http://127.0.0.1:8080`                          |
| `--username`       | 用户名                                                                       |
| `--password`       | 密码                                                                         |
| `--out`            | 输出 Excel 文件名（默认 `qb_tracker_issues.xlsx`）                           |
| `--dry-run`        | 预览模式：只导出 Excel，不实际打标签                                         |
| `--clear-old`      | 打标签前先移除本脚本此前打的所有 `错误/*` 标签                               |
| `--multi-tag`      | 一个种子命中多类时，全部打上对应标签（默认仅最高优先级）                     |
| `--progress-step`  | 每扫描多少种子打印一次进度（默认 500），数量大时可调大以减少刷屏             |

## 使用示例

**1. 预览模式 - 仅导出 Excel，不打标签**

```bash
python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --username admin --password 123456 --out qb_tracker_issues.xlsx --dry-run
```

**2. 执行模式 - 打标签（先清理旧标签再分类）**

```bash
python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --username admin --password 123456 --out qb_tracker_issues.xlsx --clear-old
```

**3. 多标签模式 - 命中多类全部打上**

```bash
python qbittorrent-tracker-tagger.py --host http://127.0.0.1:8080 --username admin --password 123456 --clear-old --multi-tag
```

## 一键运行

双击 `qbittorrent-tracker-tagger.bat` 即可（连接参数已在 bat 内配置，默认执行模式带 `--clear-old`）。

## 导出字段说明

`qb_tracker_issues.xlsx` 的 `issues` 工作表包含以下字段：

| 字段            | 说明                       |
| --------------- | -------------------------- |
| `name`          | 种子名称                   |
| `hash`          | 种子哈希                   |
| `tracker_url`   | 报错的 tracker 地址        |
| `tracker_level` | `错误` 或 `警告`            |
| `tracker_msg`   | tracker 返回的原始报错文案 |
| `matched_tag`   | 命中的标签                 |

`summary` 工作表汇总各标签命中的种子数量。

## 注意事项

1. **只读安全**：脚本不会删除任何种子或文件，仅添加标签
2. **建议先预览**：首次使用加 `--dry-run` 确认分类无误后再正式打标签
3. **重跑用 `--clear-old`**：重复执行可避免标签重复堆积
4. **API 兼容性**：需要 qBittorrent 开启 Web API

---

## License

MIT
