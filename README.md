# 财富广场任务管理 v1.0

中信「统一运营管理平台」（财富广场）营销任务的**批量创建工具**。用 Excel 做数据源，自动在后台完成「新建任务 → 选择奖励 → 填名称/链接 → 选任务类型 → 上传图片 → 选链接类型/渠道 → 提交审批」全流程。

- 行数不限：表格有多少行就跑多少行
- 断点续跑：提交过的自动跳过，中断后重跑同一条命令即可
- 失败可控：单条失败留截图 + 日志，连续 3 条失败自动熔断

> ⚠️ 内部工具，涉及公司内网地址与业务数据，请勿公开。

---

## 1. 环境要求

| 项 | 说明 |
|---|---|
| Python | 3.8+，**无需第三方库**（xlsx 用标准库 zipfile 解析） |
| BrowserSkill | `bsk` CLI ≥ 0.5 + Chrome 扩展已连接（`bsk status` 能看到浏览器） |
| 网络 | 能访问 `wealthplaza.cfhc.citic` 与 `iaas.intra.cfhc.citic`（VPN 已连） |
| 账号 | 已在后台登录，能看到「活动管理工具 → 任务管理」 |

---

## 2. 快速开始

### 2.1 准备表格

xlsx，单个工作表，**第 1 行就是数据，不要表头**。

| A（任务名称） | B（跳转链接） | C（链接所属渠道） |
|---|---|---|
| 来逛中信财富节 | https://... | 中信财富广场公众号 |

可选列：

- **图片列**：填图片文件名，相对路径按 Excel 所在目录解析（如 `pic/001.png`），配 `col_image`
- **任务组列**：每行不同的任务组，配 `col_group`

### 2.2 改配置 `config.json`

| 键 | 含义 |
|---|---|
| `excel_path` / `sheet_name` | 数据源 |
| `data_start_row` | 起始行（默认 1） |
| `col_name` / `col_url` / `col_channel` | 列字母 |
| `col_image` / `col_group` | 可选，留 `null` 表示用固定值 |
| `image_path` | 统一使用的图片（无图片列时） |
| `reward_name` / `reward_pick_nth` | 任务奖励文案；同名多项时取第几个（从 0 数） |
| `task_group` / `task_type` / `link_type` | 三个固定下拉值 |
| `state_path` | 断点记录位置 |
| `skip_invalid_rows` | 缺字段的行是否跳过（默认 true） |
| `stop_on_error` | 单条失败是否中止（默认 true，建议配合 `--continue-on-error`） |

### 2.3 起图片服务（上传必需）

```bash
python tools/asset_server.py          # 默认服务 assets/，端口 8765
python tools/asset_server.py D:/pics 9000
```

脚本运行期间必须保持开启。原理：浏览器原生文件注入被 CDP 拒绝，改为页面内 `fetch` 本地图片 → 构造 File → 派发 `change`。

### 2.4 开浏览器会话

```bash
bsk open "https://wealthplaza.cfhc.citic/lm-activity/cross-selling-backend-front/index.html#/mktSystem/taskManagement/index"
# 记下输出里的 session=xxxx

# 若新窗口停在登录页，借用已登录的标签页：
bsk tab list
bsk tab borrow <标签页id> --session xxxx
```

### 2.5 三段式执行

```bash
python batch_tasks.py --list --session xxxx              # ① 预检：只看数据，不碰浏览器
python batch_tasks.py --only 1 --dry-run --session xxxx  # ② 单条试跑：填完停在提交前，人工核对
python batch_tasks.py --session xxxx                     # ③ 全量
```

---

## 3. 命令行参数

| 参数 | 作用 |
|---|---|
| `--excel` / `--sheet` / `--session` / `--image` | 覆盖配置 |
| `--from N` / `--to N` | 只跑第 N..M 条 |
| `--only N` | 只跑第 N 条（配合 `--dry-run` 最常用） |
| `--limit N` | 本次最多跑 N 条 |
| `--list` | 只打印数据，不碰浏览器 |
| `--dry-run` | 填完不提交 |
| `--force` | 忽略断点记录强制重跑 |
| `--reset-state` | 清空断点记录 |
| `--continue-on-error` | 单条失败继续（连续 3 条失败仍会中止） |

---

## 4. 运行产物

```
runs/run_<时间戳>.log        每轮完整日志
runs/fail_n<N>_*.png         失败现场截图
runs/dryrun_n<N>_*.png       试跑截图
runs/report_<时间戳>.json    成功/失败/未处理编号
state/submitted.json         断点记录
```

---

## 5. 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| `ERR_CONNECTION_CLOSED`、页面打不开 | VPN / 内网断开；恢复后从门户重开带 `code` 的链接重新登录 |
| 下拉选不中 | 配置里的文案与页面选项不完全一致（全角/半角、空格、括号） |
| 找不到「链接类型」 | 该字段仅在「任务类型 = 普通任务-浏览」时出现，检查 `task_type` |
| 提交后判定失败但实际成功 | 历史问题，已改为「网络是否新增 `commitNewMissionEdit` 请求」判定，不会误判 |
| 长链接报「填写未生效」 | 已修复：长文本改用 DOM 真值校验，不经 VOM |
| 重复提交 | 断点记录被绕过（用了 `--force`）或手工改过任务名 |
| 换活动后没跑任何行 | 上一个活动的断点记录还在，执行 `--reset-state` |

---

## 6. 版本记录

**v1.0**
- 支持任意行数、断点续跑、分批执行
- 每行独立图片 / 任务组
- 提交结果以网络请求为证据，避免误判与重复提交
- 长链接走 DOM 真值校验
- 失败留证 + 连续失败熔断

---

## 7. 目录结构

```
batch_tasks.py          主脚本
config.json             配置
parse_xlsx.py           零依赖 xlsx 解析（导出 TSV，便于排查数据）
tools/asset_server.py   本地上传图服务（带 CORS）
assets/                 待上传图片
state/                  断点记录
runs/                   日志与截图
```
