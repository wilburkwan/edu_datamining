import pandas as pd
import matplotlib
matplotlib.use('Agg') # 設定非互動式後端，必須在 pyplot 匯入前
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from scipy import stats
import sys
import csv
import statistics
from matplotlib.font_manager import fontManager # 用於字體設定
import logging
import traceback

# 設定中文字體，以便在圖表中正確顯示中文
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS'] # 微軟正黑體, 黑體, Unicode字體
plt.rcParams['axes.unicode_minus'] = False  # 解決負號顯示問題

def setup_chinese_font():
    """
    嘗試為 Matplotlib 設定中文字型。
    """
    try:
        # 不同作業系統上常見的中文字型
        font_preferences = [
            'Microsoft JhengHei UI', # Windows (繁體)
            'Microsoft JhengHei',    # Windows (繁體)
            'PingFang TC',           # macOS (繁體)
            'Heiti TC',              # macOS (繁體)
            'SimHei',                # 簡體中文，但許多系統有
            'WenQuanYi Zen Hei',     # Linux
            'AR PL UKai CN'          # Linux
        ]
        
        available_fonts = {font.name for font in fontManager.ttflist}
        
        for font_name in font_preferences:
            if font_name in available_fonts:
                plt.rcParams['font.sans-serif'] = [font_name]
                plt.rcParams['axes.unicode_minus'] = False # 正確顯示負號
                print(f"已設定中文字型為: {font_name}")
                return
        
        print("警告：未找到偏好的中文字型。圖表中的中文可能無法正確顯示。")
        print("請嘗試安裝 'Microsoft JhengHei UI' (Windows), 'PingFang TC' (macOS) 或 'WenQuanYi Zen Hei' (Linux) 等字型。")
        plt.rcParams['font.sans-serif'] = ['sans-serif'] # 預設回退
        plt.rcParams['axes.unicode_minus'] = False

    except Exception as e:
        print(f"設定中文字型時發生錯誤: {e}")
        plt.rcParams['font.sans-serif'] = ['sans-serif'] # 預設回退
        plt.rcParams['axes.unicode_minus'] = False

def process_files(scores_file_path, class_data_file_path):
    """
    處理成績檔案和班級對照表檔案。
    主要功能：
    1. 讀取成績資料。
    2. (已移除) 讀取班級對照表，並試圖找出金鼎國小的學校代碼。
    3. (已移除) 合併資料（如果需要）。
    4. 清理和預處理資料。
    """
    df_scores = None
    identified_kinding_code = None # 雖然不再從班級檔獲取，但保留變數以防未來使用或來自成績檔

    # 1. 讀取成績資料
    try:
        # 使用 cp950 編碼讀取 CSV 檔案
        df_scores = pd.read_csv(scores_file_path, encoding='cp950')
        print(f"成功讀取成績檔案: {scores_file_path}")

        # 顯示所有現有欄位以便診斷
        print("CSV檔案的現有欄位:")
        for i, col in enumerate(df_scores.columns):
            print(f"{i}: {col}")
        
        # 直接映射現有欄位到我們需要的欄位名稱
        column_mapping = {
            '年度': '學年度',
            '縣市': '縣市',
            '學校代碼': '學校代碼',
            '學校名稱': '學校名稱',
            '年級': '年級',
            '班級': '班級代號_成績檔',
            '總平均': '整體答對率'
        }
        
        # 應用映射
        df_scores_mapped = df_scores.rename(columns=column_mapping)
        
        # 添加 學校名稱_原始 欄位 (用於熱力圖)
        df_scores_mapped['學校名稱_原始'] = df_scores_mapped['學校名稱']
        
        # 檢查必要欄位是否存在
        required_columns = ['學校代碼', '學校名稱', '年級', '班級代號_成績檔', '整體答對率']
        missing_columns = [col for col in required_columns if col not in df_scores_mapped.columns]
        
        if missing_columns:
            print(f"警告：映射後仍缺少必要欄位 {missing_columns}")
            print("映射後的欄位:")
            for col in df_scores_mapped.columns:
                print(f"- {col}")
            
            return None, None
        else:
            df_scores = df_scores_mapped
            print("欄位映射成功")

        # 轉換資料型態
        df_scores['整體答對率'] = pd.to_numeric(df_scores['整體答對率'], errors='coerce')
        df_scores['年級'] = pd.to_numeric(df_scores['年級'], errors='coerce')
        df_scores['學校代碼'] = df_scores['學校代碼'].astype(str)
        df_scores['班級代號_成績檔'] = df_scores['班級代號_成績檔'].astype(str)

        # 移除必要欄位為 NaN 的資料列
        df_scores.dropna(subset=['整體答對率', '年級', '學校代碼', '班級代號_成績檔'], inplace=True)
        
        if not df_scores.empty and '整體答對率' in df_scores.columns:
             print(f"成功讀取並初步處理成績檔案: {scores_file_path.split('/')[-1]} (共 {len(df_scores)} 筆有效資料)")
             print(f"檔案 {scores_file_path.split('/')[-1]} 的平均整體答對率: {df_scores['整體答對率'].mean():.2f}")

        # 讀取班級資訊 CSV (縣立金鼎國小_班級.csv)
        if class_data_file_path is not None:
            try:
                df_class_info = pd.read_csv(class_data_file_path, header=[0, 1], encoding='big5')
                print(f"成功讀取班級資訊檔案: {class_data_file_path.split('/')[-1]} (共 {len(df_class_info)} 筆資料)")
                
                # 嘗試從 df_class_info 提取金鼎國小學校代碼
                if not df_class_info.empty:
                    # 檢查可能的學校代碼欄位名稱 (考慮 MultiIndex)
                    potential_code_columns = [('學校資訊', '學校代碼'), ('基本資料', '學校代碼'), '學校代碼']
                    found_col = None
                    for col_candidate in potential_code_columns:
                        if col_candidate in df_class_info.columns:
                            found_col = col_candidate
                            break
                    
                    if found_col:
                        codes = df_class_info[found_col].astype(str).unique()
                        if len(codes) > 0:
                            identified_kinding_code = codes[0] # 假設第一個是金鼎國小的代碼
                            print(f"從班級資訊檔案中識別出學校代碼為: {identified_kinding_code} (將用於識別金鼎國小)")
                            if len(codes) > 1:
                                print(f"警告：班級資訊檔案中找到多個學校代碼: {codes}。已使用第一個: {identified_kinding_code}")
                        else:
                            print("警告：班級資訊檔案的學校代碼欄位中未找到有效代碼。")
                    else:
                        print(f"警告：在班級資訊檔案中未能定位到學校代碼欄位。檢查的欄位包括: {potential_code_columns}")

            except UnicodeDecodeError:
                print(f"使用 Big5 讀取班級資訊檔案 {class_data_file_path.split('/')[-1]} 失敗，嘗試 UTF-8...")
                try:
                    df_class_info = pd.read_csv(class_data_file_path, header=[0, 1], encoding='utf-8')
                    print(f"成功使用 UTF-8 讀取班級資訊檔案: {class_data_file_path.split('/')[-1]} (共 {len(df_class_info)} 筆資料)")
                    # (此處可重複提取學校代碼的邏輯)
                except Exception as e_utf8:
                    print(f"使用 UTF-8 讀取班級資訊檔案 {class_data_file_path.split('/')[-1]} 也失敗: {e_utf8}")
            except Exception as e_class:
                print(f"讀取或處理班級資訊檔案 {class_data_file_path.split('/')[-1]} 時發生錯誤: {e_class}")
        else:
            print("未提供班級資訊檔案路徑，將跳過從班級檔案中識別學校代碼的步驟。")

    except FileNotFoundError as e:
        print(f"錯誤：找不到檔案 {e.filename}")
        return None, None
    except pd.errors.EmptyDataError:
        print(f"錯誤：檔案 {scores_file_path} 為空。")
        return None, None
    except Exception as e_scores:
        print(f"處理成績檔案 {scores_file_path} 時發生未預期錯誤: {e_scores}")
        return None, None
        
    return df_scores, identified_kinding_code

def generate_comparison_chart(df_scores, kinmen_school_code, target_class_name, output_image_path):
    """
    產生比較金門縣七年級各校與特定班級數學成績分佈的圖表。
    """
    setup_chinese_font() # 在繪圖前設定字型

    if df_scores is None or df_scores.empty:
        print("錯誤：未提供有效的成績資料，無法產生圖表。")
        return

    # 確保 '整體答對率' 是數值型態
    if '整體答對率' not in df_scores.columns:
        print("錯誤：成績資料中缺少 '整體答對率' 欄位。")
        return
    df_scores['整體答對率'] = pd.to_numeric(df_scores['整體答對率'], errors='coerce')

    # 篩選七年級資料 (雖然檔名已指明，但作為雙重確認和彈性)
    # 假設 '年級' 欄位存在且為數值或可轉換為數值
    if '年級' in df_scores.columns:
        df_scores['年級'] = pd.to_numeric(df_scores['年級'], errors='coerce')
        df_grade_filtered = df_scores[df_scores['年級'] == 7].copy() # 使用 .copy() 避免 SettingWithCopyWarning
    else:
        print("警告：成績資料中缺少 '年級' 欄位，將假設所有資料均為七年級。")
        df_grade_filtered = df_scores.copy()

    if df_grade_filtered.empty:
        print("找不到七年級的成績資料。")
        return

    # 1. 金鼎國小目標班級的成績 (此處的 "金鼎國小" 僅為範例，實際應根據您的需求調整)
    target_school_scores = pd.Series(dtype=float)
    can_compare_target_school = False
    if kinmen_school_code and target_class_name:
        # 確保學校代碼和班級代號欄位存在且型態正確
        if '學校代碼' in df_grade_filtered.columns and '班級代號_成績檔' in df_grade_filtered.columns:
            df_grade_filtered['學校代碼'] = df_grade_filtered['學校代碼'].astype(str)
            df_grade_filtered['班級代號_成績檔'] = df_grade_filtered['班級代號_成績檔'].astype(str)
            
            target_school_df = df_grade_filtered[
                (df_grade_filtered['學校代碼'] == str(kinmen_school_code)) &
                (df_grade_filtered['班級代號_成績檔'] == str(target_class_name))
            ]
            target_school_scores = target_school_df['整體答對率'].dropna()
            
            if not target_school_scores.empty:
                print(f"學校 {kinmen_school_code} {target_class_name}班 學生人數: {len(target_school_scores)}, 平均答對率: {target_school_scores.mean():.2f}")
                can_compare_target_school = True
            else:
                print(f"在資料中找不到學校 {kinmen_school_code} {target_class_name}班 的有效成績。")
        else:
            print("警告：缺少 '學校代碼' 或 '班級代號_成績檔' 欄位，無法篩選目標學校班級。")
    else:
        print("未提供目標學校代碼或目標班級名稱，跳過特定班級比較。")
       
    # 2. 篩選出其他學校的七年級資料
    if kinmen_school_code and '學校代碼' in df_grade_filtered.columns:
        df_other_schools_grade_filtered = df_grade_filtered[df_grade_filtered['學校代碼'] != str(kinmen_school_code)]
    else:
        # 如果沒有目標學校代碼，則所有學校都視為 "其他學校"
        df_other_schools_grade_filtered = df_grade_filtered.copy()
        if not kinmen_school_code and '學校代碼' in df_grade_filtered.columns:
            print("警告：未指定目標學校代碼，'其他學校' 將包含所有學校的資料。")


    # 3. 計算其他學校七年級所有學生的整體答對率
    other_schools_all_students_scores = df_other_schools_grade_filtered['整體答對率'].dropna()

    if not other_schools_all_students_scores.empty:
        print(f"其他學校七年級學生總人數: {len(other_schools_all_students_scores)}, 平均答對率: {other_schools_all_students_scores.mean():.2f}")
    else:
        print("找不到其他學校七年級的成績資料。")

    # 4. 計算其他學校各班級的平均答對率
    other_schools_class_avg_scores = pd.Series(dtype=float)
    if not df_other_schools_grade_filtered.empty and '學校代碼' in df_other_schools_grade_filtered.columns and '班級代號_成績檔' in df_other_schools_grade_filtered.columns:
        other_schools_class_avg_scores = df_other_schools_grade_filtered.groupby(
            ['學校代碼', '學校名稱_原始', '班級代號_成績檔']
        )['整體答對率'].mean().dropna()
    
    if not other_schools_class_avg_scores.empty:
        print(f"其他學校七年級班級數量 (有成績者): {len(other_schools_class_avg_scores)}, 這些班級平均答對率的均值: {other_schools_class_avg_scores.mean():.2f}")
    else:
        print("找不到其他學校各班級的平均成績資料。")

    # 5. 繪圖 (使用 KDE plots)
    plt.figure(figsize=(12, 7))
    
    plot_title_parts = ["金門縣七年級成績分佈"]
    
    if can_compare_target_school:
        sns.kdeplot(target_school_scores, label=f'學校 {kinmen_school_code} {target_class_name}班 (學生)', fill=True, alpha=0.5, linewidth=1.5, color='red')
        # 特別標示目標班級的平均值
        target_mean_score = target_school_scores.mean()
        plt.axvline(target_mean_score, color='red', linestyle='-.', linewidth=1.5, 
                    label=f'目標班級平均: {target_mean_score:.2f}')
        plot_title_parts.append(f"與學校 {kinmen_school_code} {target_class_name}班比較")
        
    if not other_schools_all_students_scores.empty:
        sns.kdeplot(other_schools_all_students_scores, label='其他學校七年級 (所有學生)', fill=True, alpha=0.4, linewidth=1.5, color='skyblue')

    if not other_schools_class_avg_scores.empty:
        sns.kdeplot(other_schools_class_avg_scores, label='其他學校七年級 (班級平均分)', fill=True, alpha=0.3, linestyle='--', linewidth=1.5, color='green')

    plt.title(' '.join(plot_title_parts))
    plt.xlabel('整體答對率 (0.0 ~ 1.0)')
    plt.ylabel('密度')
    plt.xlim(0, 1) # 答對率通常在 0 到 1 之間
    plt.legend()
    plt.grid(True, linestyle=':', alpha=0.7)
    
    try:
        plt.savefig(output_image_path)
        print(f"圖表已儲存至: {output_image_path}")
    except Exception as e_save:
        print(f"儲存圖表失敗: {e_save}")
    
    # plt.show() # 如果在腳本執行時也想直接顯示圖表，取消此行註解

def generate_class_heatmap(df_scores, output_file_path):
    """
    產生全縣各班級平均成績熱力圖。
    """
    try:
        # 檢查必要欄位是否存在
        required_columns = ['整體答對率', '學校名稱_原始', '班級代號_成績檔']
        missing_columns = [col for col in required_columns if col not in df_scores.columns]
        if missing_columns:
            print(f"錯誤：成績資料中缺少必要欄位 (需要 {missing_columns})，無法產生熱力圖。")
            return

        # 將七年級改為八年級
        # 過濾出七年級的資料 (假設 '年級' 欄位存在且七年級編碼為 7)
        if '年級' in df_scores.columns:
            # 七年級 7 改為 八年級 8
            df_7th = df_scores[df_scores['年級'] == 8].copy()
            if df_7th.empty:
                print("找不到八年級的成績資料可供熱力圖分析。")
                return
        else:
            print("錯誤：成績資料中缺少 '年級' 欄位，無法過濾出八年級資料。")
            return

        # 計算各班級平均答對率
        class_avg_scores = df_7th.groupby(
            ['學校名稱_原始', '班級代號_成績檔']
        )['整體答對率'].mean().dropna()

        if class_avg_scores.empty:
            print("計算後沒有可用的班級平均成績資料可供熱力圖分析。")
            return

        # 將 Series 轉換為適合熱力圖的 DataFrame (學校為 index, 班級為 columns)
        try:
            heatmap_data = class_avg_scores.unstack(level='班級代號_成績檔')
        except Exception as e_unstack:
            print(f"將資料轉換為熱力圖格式時發生錯誤: {e_unstack}")
            # 如果 unstack 因為索引問題失敗，可以嘗試重設索引再 pivot
            try:
                class_avg_scores_df = class_avg_scores.reset_index()
                heatmap_data = class_avg_scores_df.pivot_table(
                    index='學校名稱_原始', 
                    columns='班級代號_成績檔', 
                    values='整體答對率'
                )
            except Exception as e_pivot:
                print(f"嘗試使用 pivot_table 轉換資料也失敗: {e_pivot}")
                return


        if heatmap_data.empty or heatmap_data.isnull().all().all():
            print("產生的熱力圖資料為空或全為 NaN。")
            return

        # 繪製熱力圖
        # 動態調整圖表大小以容納所有標籤
        # 欄寬至少為1，行高至少為0.6
        fig_width = max(12, len(heatmap_data.columns) * 1.2) 
        fig_height = max(8, len(heatmap_data.index) * 0.7)
        
        plt.figure(figsize=(fig_width, fig_height))
        sns.heatmap(
            heatmap_data, 
            annot=True,       # 在儲存格中顯示數值
            fmt=".2f",        # 數值格式化為小數點後兩位
            cmap="YlGnBu",    # 使用 Yellow-Green-Blue 色彩對應
            linewidths=.5,    # 儲存格之間的線條寬度
            cbar_kws={'label': '平均整體答對率'} # 色條標籤
        )
        plt.title('金門縣八年級各班級平均整體答對率熱力圖', fontsize=16)
        plt.xlabel('班級代號', fontsize=12)
        plt.ylabel('學校名稱', fontsize=12)
        plt.xticks(rotation=45, ha='right') # X軸標籤旋轉以防重疊
        plt.yticks(rotation=0)
        plt.tight_layout() # 自動調整子圖參數以給定一個緊密的佈局

        try:
            plt.savefig(output_file_path)
            print(f"熱力圖已儲存至: {output_file_path}")
        except Exception as e_save:
            print(f"儲存熱力圖失敗: {e_save}")

    except Exception as e:
        print(f"產生熱力圖時發生錯誤: {e}")

def read_score_file(file_path):
    try:
        # 嘗試不同的編碼方式
        encodings = ['utf-8', 'big5', 'big5hkscs', 'gb18030', 'gbk', 'cp950']
        
        for encoding in encodings:
            try:
                print(f"嘗試使用 {encoding} 編碼讀取檔案...")
                # 讀取CSV檔案，使用當前測試的編碼
                df = pd.read_csv(file_path, encoding=encoding)
                
                # 檢查是否成功讀取並顯示前幾筆資料以確認
                print("檔案讀取成功，檢查欄位名稱:")
                print(df.columns.tolist())
                
                # 檢查必要欄位是否存在
                required_columns = ['學校代碼', '學校名稱', '年級', '班級代號_成績檔', '整體答對率']
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    print(f"警告：缺少欄位 {missing_columns}，但將嘗試繼續處理...")
                    
                    # 嘗試尋找相似欄位並重命名
                    # 顯示現有欄位以幫助診斷
                    print("現有欄位:")
                    for i, col in enumerate(df.columns):
                        print(f"{i}: {col}")
                        
                    # 在這裡可以添加欄位映射邏輯
                    # 例如：df = df.rename(columns={'現有欄位名': '需要的欄位名'})
                    
                return df
                
            except UnicodeDecodeError:
                print(f"{encoding} 編碼無法正確解析檔案，嘗試下一種編碼...")
            except Exception as e:
                print(f"使用 {encoding} 編碼時發生錯誤: {str(e)}")
        
        # 如果所有編碼都失敗
        raise Exception("無法以任何已知編碼正確讀取檔案。")
        
    except Exception as e:
        print(f"讀取成績檔案時發生錯誤: {str(e)}")
        return None

if __name__ == "__main__":
    # 設定 logging 模組
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(filename)s:%(lineno)d: %(message)s')

    # 檔案路徑設定
    scores_data_file = "113年度_學力測驗_金門縣_數學8年級成績_202406.csv" # 已正確使用八年級成績檔
    # 注意：以下班級資訊檔案和目標班級名稱是針對四年級的。
    # 如果您也需要精確的三年級特定班級比較圖，請更新這些路徑和名稱。
    # 對於全縣三年級熱力圖，這些特定設定的影響較小。
    # class_data_file = "康軒平台資料/第二版_113-9_114-2/縣立金鼎國小_班級.csv" # <--- 不再使用此檔案
    output_chart_file = "金門縣8年級數學成績分佈比較圖.png" # 已正確更新為八年級
    output_heatmap_file = "金門縣8年級各班平均成績熱力圖.png" # 已正確更新為八年級
    
    # 金鼎國小的目標班級名稱 (目前為四年級，對三年級熱力圖非必需，但影響比較圖)
    # TARGET_CLASS_NAME_MAIN = '302' # <--- 不再使用，因為 class_data_file 已被忽略

    print("開始處理檔案並產生圖表...")
    # 設定中文字型
    setup_chinese_font() # 確保呼叫字型設定函數

    # 修改 process_files 的調用，不再傳遞 class_data_file 路徑
    # identified_kinding_code 將不會從班級檔案中獲取
    df_processed_scores, identified_kinding_code = process_files(scores_data_file, None) # 傳遞 None 給 class_data_file 參數
    
    if df_processed_scores is not None and not df_processed_scores.empty:
        # 由於不再處理 class_data_file，identified_kinding_code 主要來自成績檔案（如果有的話）或為 None
        # 相關的提示訊息已移除，因為比較圖功能已註解

        # 產生比較圖 (請注意上述關於目標班級和學校代碼的提示)
        # 如果您只想產生熱力圖，可以註解掉下一行
        # generate_comparison_chart(df_processed_scores, identified_kinding_code, TARGET_CLASS_NAME_MAIN, output_chart_file)
        
        # 產生八年級熱力圖
        generate_class_heatmap(df_processed_scores, output_heatmap_file)
    else:
        print("由於讀取或處理成績資料失敗，無法產生圖表。")

    print("圖表產生流程結束。")
