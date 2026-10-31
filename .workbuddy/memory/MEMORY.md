# 项目长期记忆：qbittorrent-skip-hash-check

## 环境
- qBittorrent：`http://192.168.10.111:9091`，admin，**v5.3.0**，库中约 4.1 万种子。
- 本机 Python：系统 Python313（`C:\Users\Schal\AppData\Local\Programs\Python\Python313\python.exe`，双击 bat 走它）
  + WorkBuddy 托管 3.13.12/3.14.3。三个解释器都已装 requests/pandas/openpyxl。

## 鉴权（qb_auth.py，两脚本共享）
- **两种鉴权**：API token（>=5.2.0）优先于用户名密码，互斥。
- **必须 `Authorization: Bearer <token>`**。实测 v5.3.0 上 `X-API-Key` 头返回 403。
- **token 不能用于 `auth/` 端点**（login/logout 返回 403），故验证 token 改为探测 `/api/v2/app/version`（响应仅版本串，成本最低）。
- token 格式：`qbt_` + 28 位，共 32 字符。
- 密码登录兼容：旧版 `200 + "Ok."`，4.6+ `204` 空响应 → 判定用 2xx + `QBT_SID*` Cookie 二次校验。
- 密码模式结束主动 logout；token 模式无状态无需处理。

## 已知坑
- 登录判定曾写死 `status==200 and 'Ok.' in text`，在返回 204 的新版 qb 上误判"登录失败"（2026-11-01 修复）。
- Windows 控制台 GBK，脚本打印 emoji 必须 `sys.stdout.reconfigure(encoding='utf-8', errors='replace')`。
- bat 双击运行窗口会秒关，需 `pause` + `errorlevel` 提示。
- 大量种子时 `/torrents/info` 响应数十 MB，读取超时放宽到 60s（连接 10s）。

## 性能结论
token 对**单次脚本运行**收益可忽略（省 1 次登录往返 + 服务端 PBKDF2 约 5–50ms，局域网总省 10–100ms）；
真正受益的是高频轮询/常驻集成，以及免会话堆积、免失败封 IP。

## 文件
- `qb_auth.py` 共享鉴权（新增，2026-11-01）
- `qbittorrent-skip-hash-check.py/.bat` 跳过校验（删除后重导入，skip_checking）
- `qbittorrent-tracker-tagger.py/.bat` 按 tracker 错误分类打标签，7 类，只打标签不删除
