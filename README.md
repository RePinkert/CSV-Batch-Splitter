# CSV Batch Splitter

一个简单高效的CSV文件分批处理工具，能够将大型CSV文件按指定规则分割成小批次，并生成Excel文件。

## ✨ 主要功能

- 🔄 **智能分批处理**：自动将大型CSV文件分割成小批次
- 🔁 **循环切分支持**：当接近文件末尾时自动从表头循环继续，确保每次都能生成完整批次
- 📊 **Excel导出**：生成的批次自动保存为Excel格式
- 📝 **详细日志**：完整记录每天的批次生成情况，循环切分会记录为子批次
- ⚙️ **灵活配置**：可自定义批次大小、每日限制等参数

## 🛠️ 环境要求

- **Python 3.7+**（推荐3.8+）
- **必需依赖**：`pandas` 和 `openpyxl`

## 🚀 使用方法

### 0️⃣ 获取项目文件
```bash
# 使用Git克隆（推荐）
git clone <repository-url>
cd csv_slide

# 或直接下载项目文件到本地目录
```

### 1️⃣ 安装依赖
```bash
pip install pandas openpyxl
```

### 2️⃣ 放置 CSV 文件
将你的待处理csv文件（例如 `ue.csv`）放在和 `test.py` 同一目录下即可。

示例格式：
```csv
email
aaa@example.com
bbb@example.com
...
```

### 3️⃣ 执行脚本
```bash
python test.py
```

运行后会生成：
```
mail_batch_2025-01-01_b1.xlsx
```

## ⚙️ 配置说明（来自 test.py）

```python
MAX_RECIPIENTS_PER_EMAIL = 200   # 每批 200 人
MAX_EMAILS_PER_DAY = 20          # 每天最多 20 批
HAS_HEADER = True                # CSV 是否带表头
CSV_PATH = "ue.csv"              # 未找到时自动搜索 *.csv
```

## 🔁 进度管理（state.json）

示例：
```json
{
  "current_index": 38
}
```

**含义**：上次已处理到第 38 行（从 0 开始计，不含表头）。

- 当到达文件末尾 → 自动归零，从头循环切分
- 不需要人工修改

## 📝 日志格式（send_log.csv）

示例：
```csv
date,batch_no,count,excel_start_row,excel_end_row
2025-01-01,1,200,2,201
2025-01-01,2,200,202,401
2025-01-01,3-1,162,15986,16147
2025-01-01,3-2,38,2,39
```

**说明**：
- `3-1` / `3-2`：一次切分不足 200 行，自动从表头补足
- `excel_start_row` / `excel_end_row`：对应 Excel 的真实行号（从 2 行开始，因为第 1 行是标题）

## 🔄 循环切分示例

当 CSV 只剩 162 行时：

```
🔄 剩余 162 行，不足 200 行
→ 从表头补足 38 行
→ 本批共 200 行
→ 日志写入 2 条（例如 3-1 与 3-2）
```

生成的 Excel 依然包含 200 行完整批次。

## 🧪 示例流程

假设 CSV 有 20,000 行：
- 执行 1 次 → 批次1（200 行）
- 执行 20 次 → 当天批次耗尽
- 第二天继续 → 批次21
- 直至 20,000 行被循环切分完毕

## ⚠️ 注意事项

- 每次执行脚本 只生成 1 个 Excel 文件（即用即删）
- 每天最多生成 20 个的限制是可以修改的
- CSV 必须为 UTF-8 编码，否则可能出现乱码

## 🆘 故障排查

**❗ModuleNotFoundError: No module named 'pandas'**
安装依赖：
```bash
pip install pandas openpyxl
```

**❗ 找不到 CSV 文件**
- 请确保脚本目录下至少有一个 `.csv` 文件
- 或检查文件路径是否正确命名为 `ue.csv`
- 或修改 `CSV_PATH` 变量

## 📊 输出文件说明

| 文件类型 | 说明 |
|---------|------|
| `mail_batch_YYYY-MM-DD_bN.xlsx` | 生成的批次Excel文件 |
| `send_log.csv` | 批次生成日志 |
| `state.json` | 处理进度记录 |

---

**提示**: 如果这个工具对您有帮助，欢迎给项目点个星⭐！