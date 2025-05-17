import csv
import os

def extract_school_name_from_filepath(filepath):
    """從檔案路徑中提取學校名稱。"""
    # 例如: "康軒平台資料/第二版_113-9_114-2/縣立中正國中_班級.csv" -> "縣立中正國中"
    filename = os.path.basename(filepath)
    school_name_part = filename.replace("_班級.csv", "")
    return school_name_part

def parse_csv_for_math_stats(filepath):
    """解析單個CSV檔案以獲取數學測驗統計數據。"""
    school_name = extract_school_name_from_filepath(filepath)
    all_rows = []
    
    try:
        with open(filepath, 'r', encoding='utf-8-sig') as f: # utf-8-sig 處理 BOM
            reader = csv.reader(f)
            all_rows = list(reader)
    except FileNotFoundError:
        print(f"錯誤：找不到檔案 {filepath}")
        return school_name, 0, 0, "N/A", "N/A"

    test_section_start_idx = -1
    for i, row in enumerate(all_rows):
        if row and row[0] == "測驗":
            test_section_start_idx = i
            break

    if test_section_start_idx == -1:
        print(f"警告：在檔案 {filepath} 中找不到 '測驗' 區塊。")
        return school_name, 0, 0, "N/A", "N/A"

    # 數學相關欄位索引 (在 "測驗" -> "班級,派出任務數,..." 這行之後)
    # 班級 (0), 國派 (1), ..., 英正 (8), 數派 (9), 數完 (10), 數成 (11), 數正 (12), ...
    math_tasks_assigned_col = 9
    math_tasks_completed_col = 10
    math_correct_rate_col = 12

    total_math_assigned = 0
    total_math_completed = 0
    sum_of_valid_correct_rates = 0.0
    count_of_valid_correct_rates = 0

    # 迭代 "測驗" 區塊中的資料行 (跳過標頭)
    for i in range(test_section_start_idx + 2, len(all_rows)):
        row = all_rows[i]
        
        # 檢查是否為資料結束的空行或分隔行
        if not row or not row[0] or row[0] == "班級": # 跳過空行, 無班級名稱行, 或重複的標頭行
            if all(not cell or cell.isspace() for cell in row): # 如果整行都是空的或只有空白，視為區塊結束
                 break
            continue
        
        # 確保行中有足夠的欄位
        if len(row) <= max(math_tasks_assigned_col, math_tasks_completed_col, math_correct_rate_col):
            # print(f"警告：檔案 {filepath} 中的行資料欄位不足，已跳過: {row}")
            continue

        try:
            class_assigned_str = row[math_tasks_assigned_col]
            class_completed_str = row[math_tasks_completed_col]
            class_correct_rate_str = row[math_correct_rate_col]

            class_assigned = int(class_assigned_str) if class_assigned_str.strip() else 0
            class_completed = int(class_completed_str) if class_completed_str.strip() else 0
            
            total_math_assigned += class_assigned
            total_math_completed += class_completed

            if class_completed > 0: # 只有在完成任務數大於0時，才考慮其正答率
                if class_correct_rate_str.strip(): # 確保正答率字串不是空的
                    try:
                        correct_rate = float(class_correct_rate_str)
                        sum_of_valid_correct_rates += correct_rate
                        count_of_valid_correct_rates += 1
                    except ValueError:
                        # print(f"警告：無法將 {class_correct_rate_str} 轉換為數字 (正答率)，已跳過。行: {row}")
                        pass # 跳過無法轉換的正答率
        except ValueError:
            # print(f"警告：無法轉換任務數為數字，已跳過。行: {row}")
            continue # 跳過轉換任務數出錯的行
        except IndexError:
            # print(f"警告：檔案 {filepath} 中的行資料欄位索引錯誤，已跳過: {row}")
            continue


    school_overall_completion_rate = (total_math_completed / total_math_assigned) if total_math_assigned > 0 else 0.0
    
    if count_of_valid_correct_rates > 0:
        school_avg_correct_rate_val = sum_of_valid_correct_rates / count_of_valid_correct_rates
        school_avg_correct_rate_formatted = f"{school_avg_correct_rate_val:.2%}"
    else:
        school_avg_correct_rate_formatted = "N/A"

    school_overall_completion_rate_formatted = f"{school_overall_completion_rate:.2%}"

    return school_name, total_math_assigned, total_math_completed, school_overall_completion_rate_formatted, school_avg_correct_rate_formatted

def main():
    # 獲取腳本檔案 (a.py) 所在的目錄
    # 例如，如果 a.py 在 C:\Users\will\Desktop\edu_datamining-main\final\a.py
    # 那麼 script_dir 將是 C:\Users\will\Desktop\edu_datamining-main\final
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 根據使用者提供的路徑結構，CSV 檔案位於腳本所在目錄下的 "康軒平台資料/第二版_113-9_114-2/" 子目錄中
    # 所以 base_path 應該是 script_dir 加上這個相對路徑
    # 例如: C:\Users\will\Desktop\edu_datamining-main\final\康軒平台資料\第二版_113-9_114-2
    base_path = os.path.join(script_dir, "康軒平台資料", "第二版_113-9_114-2")
    
    files_to_process = [
        os.path.join(base_path, "縣立中正國中_班級.csv"),
        os.path.join(base_path, "縣立上岐國小_班級.csv"),
        os.path.join(base_path, "縣立中正國小_班級.csv"),
        os.path.join(base_path, "縣立何浦國小_班級.csv"),
        os.path.join(base_path, "縣立卓環國小_班級.csv"),
        os.path.join(base_path, "縣立古城國小_班級.csv"),
        os.path.join(base_path, "縣立古寧國小_班級.csv"),
        os.path.join(base_path, "縣立多年國小_班級.csv"),
        os.path.join(base_path, "縣立安瀾國小_班級.csv"),
        os.path.join(base_path, "縣立柏村國小_班級.csv"),
        os.path.join(base_path, "縣立正義國小_班級.csv"),
        os.path.join(base_path, "縣立湖埔國小_班級.csv"),
        os.path.join(base_path, "縣立烈嶼國中_班級.csv"),
        os.path.join(base_path, "縣立西口國小_班級.csv"),
        os.path.join(base_path, "縣立賢庵國小_班級.csv"),
        os.path.join(base_path, "縣立金寧國中_班級.csv"),
        os.path.join(base_path, "縣立述美國小_班級.csv"),
        os.path.join(base_path, "縣立金城國中_班級.csv"),
        os.path.join(base_path, "縣立金寧國小_班級.csv"),
        os.path.join(base_path, "縣立金沙國中_班級.csv"),
        os.path.join(base_path, "縣立金沙國小_班級.csv"),
        os.path.join(base_path, "縣立金湖國中_班級.csv"),
        os.path.join(base_path, "縣立金湖國小_班級.csv"),
        os.path.join(base_path, "縣立金鼎國小_班級.csv"),
        os.path.join(base_path, "縣立開瑄國小_班級.csv"),
        os.path.join(base_path, "顯立古城國小_班級.csv") # 根據您的檔案列表，新增此檔案
    ]

    results = []
    for filepath in files_to_process:
        if not os.path.exists(filepath):
            print(f"錯誤：檔案路徑不存在 {filepath}")
            school_name = extract_school_name_from_filepath(filepath) # 嘗試提取名稱用於報告
            results.append((school_name, 0, 0, "N/A", "N/A (檔案未找到)"))
            continue
        results.append(parse_csv_for_math_stats(filepath))

    print("\n學校數學成績摘要：")
    print("-" * 80)
    print(f"{'學校名稱':<15} | {'總派出數學任務數':<15} | {'總已完成數學任務數':<18} | {'整體數學完成率':<15} | {'平均數學正答率':<15}")
    print("-" * 80)
    for res in results:
        school_name, assigned, completed, completion_rate, correct_rate = res
        print(f"{school_name:<15} | {assigned:<18} | {completed:<20} | {completion_rate:<18} | {correct_rate:<15}")
    print("-" * 80)

    print("\n說明：")
    print("1. 整體數學完成率 = 該校所有班級總已完成數學任務數 / 該校所有班級總派出數學任務數。")
    print("2. 平均數學正答率 = 對該校所有已完成數學任務數 > 0 的班級，其正答率的算術平均值。")
    print("3. 如果某項數據為 'N/A'，表示相關原始數據不足或檔案處理問題。")
    print("\n您可以將以上表格數據複製到試算表軟體（如 Excel, Google Sheets）中製作圖表，")
    print("例如比較各校的完成率或正答率。")
    print("若要使用 Python 繪圖，可以考慮使用 Matplotlib 或 Seaborn 等庫。")

if __name__ == "__main__":
    main()
