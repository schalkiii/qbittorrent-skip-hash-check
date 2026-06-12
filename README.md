# qBittorrent Skip Hash Check

跳过 qBittorrent 校验的工具，支持预览模式和实际执行模式。

## 功能特性

- **预览模式**：导出指定状态的种子信息到 Excel，便于确认
- **跳过校验**：将校验中的种子改为做种状态（通过删除后重新导入实现）
- **解除强制做种**：将强制做种的种子改为普通做种状态
- **批量处理**：支持一次性处理多个种子

## 环境要求

- Python 3.6+
- 依赖库：`requests`、`pandas`、`openpyxl`

```bash
pip install requests pandas openpyxl
```

## 使用方法

### 基本参数

| 参数            | 说明                                                |
| --------------- | --------------------------------------------------- |
| `--host`        | qBittorrent Web UI 地址，如 `http://127.0.0.1:8080` |
| `--username`    | 用户名                                              |
| `--password`    | 密码                                                |
| `--filter`      | 过滤类型：`checking`、`completed`、`forced`、`all`  |
| `--out`         | 输出 Excel 文件名                                   |
| `--skip-check`  | 跳过校验中的种子                                    |
| `--skip-forced` | 跳过强制做种的种子                                  |

### 过滤选项说明

| 选项        | 说明                 |
| ----------- | -------------------- |
| `checking`  | 校验中的种子（默认） |
| `completed` | 已完成的种子         |
| `forced`    | 强制做种的种子       |
| `all`       | 全部种子             |

### 使用示例

**1. 预览模式 - 查看校验中的种子**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password 123456 --filter checking --out checking.xlsx
```

**2. 执行模式 - 跳过校验**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password 123456 --filter checking --out checking.xlsx --skip-check
```

**3. 预览模式 - 查看强制做种的种子**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password 123456 --filter forced --out forced.xlsx
```

**4. 执行模式 - 解除强制做种**

```bash
python qbittorrent-skip-hash-check.py --host http://127.0.0.1:8080 --username admin --password 123456 --filter forced --out forced.xlsx --skip-forced
```

## 工作原理

该脚本通过以下步骤跳过种子校验：

1. **导出种子文件** - 从 qBittorrent 获取种子文件
2. **删除种子** - 删除原种子（保留文件）
3. **重新导入** - 使用 `skip_checking=true` 参数重新导入，从而跳过校验

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
2. **保留文件**：删除种子时默认保留文件，不会删除实际数据
3. **API 兼容性**：需要 qBittorrent 开启 Web API

## License

MIT
