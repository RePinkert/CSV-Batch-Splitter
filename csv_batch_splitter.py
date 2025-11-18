import csv
import json
import os
from datetime import datetime

import pandas as pd  # pip install pandas openpyxl

# ===== 配置区 =====
CSV_PATH = "ue.csv"          # 主表路径（相对于当前目录）
STATE_FILE = "state.json"    # 保存主表循环进度
LOG_FILE = "send_log.csv"    # 每日批次日志

MAX_RECIPIENTS_PER_EMAIL = 200
MAX_EMAILS_PER_DAY = 20
HAS_HEADER = True            # 主 CSV 是否有表头
# ===== 配置区结束 =====


def find_csv_file():
    """自动查找CSV文件 - 优先使用ue.csv"""
    # 首先尝试配置的路径
    if os.path.exists(CSV_PATH):
        return CSV_PATH
    
    # 如果配置的文件不存在，尝试当前目录下的CSV文件
    csv_files = [f for f in os.listdir('.') if f.lower().endswith('.csv')]
    
    # 优先查找ue.csv（源数据文件）
    ue_csv_candidates = [f for f in csv_files if f.lower() in ['ue.csv', 'ue.CSV']]
    if ue_csv_candidates:
        print(f"✅ 找到源数据文件: {ue_csv_candidates[0]}")
        return ue_csv_candidates[0]
    
    # 如果没有ue.csv，提供清晰的错误信息
    if csv_files:
        print(f"❌ 错误：未找到源数据文件 'ue.csv'")
        print(f"   当前目录发现的CSV文件: {csv_files}")
        print(f"   请将源数据文件重命名为 'ue.csv'，或删除其他CSV文件")
        return None
    
    return None

def load_csv():
    """加载主 CSV，返回 header（可能为 None）、数据行列表"""
    csv_file = find_csv_file()
    
    if not csv_file:
        print(f"❌ 错误：未找到CSV文件")
        print(f"   请确保当前目录下有CSV文件，或将CSV文件重命名为 '{CSV_PATH}'")
        return None, []
    
    try:
        with open(csv_file, newline="", encoding="utf-8") as f:
            reader = list(csv.reader(f))
    except FileNotFoundError:
        print(f"❌ 错误：找不到文件 '{csv_file}'")
        return None, []
    except Exception as e:
        print(f"❌ 错误：读取CSV文件失败 - {e}")
        return None, []
    
    if not reader:
        print(f"⚠️  警告：CSV文件 '{csv_file}' 为空")
        return None, []
    
    if HAS_HEADER:
        header = reader[0]
        rows = reader[1:]
    else:
        header = None
        rows = reader
    
    print(f"✅ 成功加载CSV文件: {csv_file}，共 {len(rows)} 条记录")
    return header, rows


def load_state():
    """加载进度：当前已经用到的数据行索引（从 0 开始）"""
    if not os.path.exists(STATE_FILE):
        return {"current_index": 0}
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def init_log_file_if_needed():
    """如果日志文件不存在，则创建并写入表头"""
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["date", "batch_no", "count", "excel_start_row", "excel_end_row"])


def get_today_batch_count():
    """统计今天已经生成了多少个Excel文件（从日志里看）"""
    if not os.path.exists(LOG_FILE):
        return 0
    today = datetime.now().strftime("%Y-%m-%d")
    batch_numbers = set()  # 使用set去重
    
    with open(LOG_FILE, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("date") == today:
                batch_no = row.get("batch_no", "")
                # 提取批次号主干（去掉"-1", "-2"后缀）
                base_batch_no = batch_no.split('-')[0]
                if base_batch_no:
                    batch_numbers.add(base_batch_no)
    
    return len(batch_numbers)


def append_log(date_str, batch_no, count, start_idx, end_idx):
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([date_str, batch_no, count, start_idx, end_idx])


def main():
    # 1. 读主表
    header, rows = load_csv()
    total = len(rows)
    print(f"主表总记录数: {total}")

    if total == 0:
        print("主 CSV 没有数据，退出。")
        return

    # 2. 读进度
    state = load_state()
    current_index = state.get("current_index", 0)

    # 如果到达末尾，则循环从头开始
    if current_index >= total:
        print("🔄 已到达表末尾，从表头开始循环切分")
        current_index = 0
        state["current_index"] = 0
        save_state(state)

    # 3. 检查今天的批次数量
    init_log_file_if_needed()
    today = datetime.now().strftime("%Y-%m-%d")
    today_batches = get_today_batch_count()
    print(f"今天已生成批次数: {today_batches}")

    if today_batches >= MAX_EMAILS_PER_DAY:
        print(f"今天的上限 {MAX_EMAILS_PER_DAY} 批已经用完，不再生成新文件。")
        return

    # 4. 检查CSV文件大小是否合理
    if total <= MAX_RECIPIENTS_PER_EMAIL:
        print(f"❌ 错误：CSV文件只有 {total} 条记录，不大于要求的每批 {MAX_RECIPIENTS_PER_EMAIL} 条")
        print(f"   请增大CSV文件内容，或减小配置文件中的 MAX_RECIPIENTS_PER_EMAIL")
        return
    
    # 5. 计算本次要切多少条（支持循环切分）
    remaining = total - current_index
    
    if remaining >= MAX_RECIPIENTS_PER_EMAIL:
        # 剩余记录足够，直接切分
        batch_size = MAX_RECIPIENTS_PER_EMAIL
        start_idx = current_index
        end_idx = current_index + batch_size
        batch_rows = rows[start_idx:end_idx]
        
        print(f"📊 正常切分：剩余记录充足，切分第 {current_index+1}-{end_idx} 条记录")
        
    else:
        # 剩余记录不足，需要循环从表头补充
        first_part = rows[current_index:]  # 从当前位置到末尾
        needed = MAX_RECIPIENTS_PER_EMAIL - len(first_part)  # 还需要多少条
        second_part = rows[:needed]  # 从表头补充
        
        batch_rows = first_part + second_part
        batch_size = len(batch_rows)
        
        # 计算显示用的索引信息
        print(f"🔄 循环切分：剩余 {remaining} 条记录不足，从表头补充 {needed} 条记录")
        print(f"   第一段：Excel行 {current_index + 1 + (1 if HAS_HEADER else 0)} ~ {total + (1 if HAS_HEADER else 0)} ({len(first_part)} 条)")
        print(f"   第二段：Excel行 {1 + (1 if HAS_HEADER else 0)} ~ {needed + (1 if HAS_HEADER else 0)} ({len(second_part)} 条)")
    
    # 统一计算结束索引（用于更新状态）
    end_idx = (current_index + batch_size) % total
    
    # 转换为Excel行号显示（更用户友好）
    if remaining >= MAX_RECIPIENTS_PER_EMAIL:
        # 正常切分的情况
        excel_start_row = current_index + 1 + (1 if HAS_HEADER else 0)
        excel_end_row = (current_index + batch_size) + (1 if HAS_HEADER else 0)
        print(f"本次将切分记录区间: Excel行 {excel_start_row} ~ {excel_end_row}, 共 {batch_size} 条")
    else:
        # 循环切分的情况，显示两段
        first_end_excel = total + (1 if HAS_HEADER else 0)
        second_end_excel = needed + (1 if HAS_HEADER else 0)
        print(f"本次将切分记录区间: Excel行 {current_index + 1 + (1 if HAS_HEADER else 0)} ~ {first_end_excel} + Excel行 {1 + (1 if HAS_HEADER else 0)} ~ {second_end_excel}, 共 {batch_size} 条")

    # 6. 生成 Excel 文件
    batch_no_today = today_batches + 1
    filename = f"mail_batch_{today}_b{batch_no_today}.xlsx"

    if header:
        df = pd.DataFrame(batch_rows, columns=header)
    else:
        df = pd.DataFrame(batch_rows)

    df.to_excel(filename, index=False)
    print(f"已生成文件: {filename}")

    # 7. 写日志 & 更新进度
    if remaining >= MAX_RECIPIENTS_PER_EMAIL:
        # 正常切分情况
        excel_log_start = current_index + (1 if HAS_HEADER else 0)
        excel_log_end = (current_index + batch_size - 1) + (1 if HAS_HEADER else 0)
        append_log(today, batch_no_today, batch_size, excel_log_start, excel_log_end)
    else:
        # 循环切分情况：记录两段信息
        print(f"📝 记录循环切分日志...")
        first_log_end = (total - 1) + (1 if HAS_HEADER else 0)
        second_log_start = 0 + (1 if HAS_HEADER else 0)
        second_log_end = (needed - 1) + (1 if HAS_HEADER else 0)
        
        # 第一段日志
        append_log(today, f"{batch_no_today}-1", len(first_part), current_index + (1 if HAS_HEADER else 0), first_log_end)
        # 第二段日志  
        append_log(today, f"{batch_no_today}-2", len(second_part), second_log_start, second_log_end)
    
    state["current_index"] = end_idx
    save_state(state)

    print(f"进度更新: current_index = {end_idx}")
    print("完成。")


if __name__ == "__main__":
    main()
