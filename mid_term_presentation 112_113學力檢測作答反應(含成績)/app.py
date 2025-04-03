import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from matplotlib.font_manager import FontProperties
import warnings
warnings.filterwarnings('ignore')

# 確保中文顯示正常
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

# 讀取資料函數 - 處理編碼問題
def read_data(file_path):
    try:
        # 嘗試使用不同的編碼方式讀取CSV
        encodings = ['utf-8', 'big5', 'gbk', 'latin1']
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                # 檢查是否有亂碼
                if '�' in str(df.columns):
                    continue
                else:
                    return df
            except:
                continue
        
        # 如果上面的方法都失敗，嘗試直接讀取而不指定編碼
        return pd.read_csv(file_path)
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return None

# 從檔名提取年級時
def extract_grade(filename):
    # 找出檔名中的數字
    grade = filename.split('數學')[1][0]  # 取出年級數字
    return int(grade)  # 轉為整數

# 讀取所有CSV文件並整合
def load_all_data(file_paths):
    all_data = []
    
    for file_path in file_paths:
        try:
            # 從檔名獲取年級信息
            grade = extract_grade(file_path)
            df = read_data(file_path)
            
            if df is not None:
                # 重命名一些關鍵列，假設這些列的位置是固定的
                # 這只是一個示例，您可能需要根據實際數據進行調整
                renamed_columns = {
                    df.columns[8]: '姓名',
                    df.columns[9]: '性別',  # 1 可能是男生，2 可能是女生
                    df.columns[2]: '學校代碼',
                    df.columns[18]: '總得分率',
                    df.columns[3]: '學校名稱'
                }
                
                # 使用位置而不是名稱重命名列
                df = df.rename(columns=renamed_columns)
                
                # 添加年級信息
                df['年級'] = grade
                
                # 清理數據
                df['總得分率'] = pd.to_numeric(df['總得分率'], errors='coerce')
                
                all_data.append(df)
        except Exception as e:
            print(f"處理檔案 {file_path} 時發生錯誤: {e}")
    
    # 將所有數據合併成一個DataFrame
    if all_data:
        combined_data = pd.concat(all_data, ignore_index=True)
        return combined_data
    else:
        return None

# 視覺化1: 不同年級的數學能力比較
def visualize_grade_differences(df):
    plt.figure(figsize=(10, 6))
    
    # 計算每個年級的平均得分率
    grade_scores = df.groupby('年級')['總得分率'].mean().reset_index()
    
    # 繪製條形圖
    sns.barplot(x='年級', y='總得分率', data=grade_scores, palette='viridis')
    plt.title('不同年級的平均數學得分率比較')
    plt.xlabel('年級')
    plt.ylabel('平均得分率')
    plt.ylim(0, 1)  # 假設得分率在0到1之間
    
    # 在每個條形上顯示具體數值
    for i, row in grade_scores.iterrows():
        plt.text(i, row['總得分率'] + 0.02, f'{row["總得分率"]:.2f}', 
                 ha='center', va='bottom')
    
    plt.savefig('不同年級數學能力比較.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    # 繪製箱型圖來顯示分布
    plt.figure(figsize=(12, 6))
    sns.boxplot(x='年級', y='總得分率', data=df, palette='viridis')
    plt.title('不同年級數學得分率分布')
    plt.xlabel('年級')
    plt.ylabel('得分率')
    plt.ylim(0, 1)
    plt.savefig('不同年級數學能力分布.png', dpi=300, bbox_inches='tight')
    plt.show()

# 視覺化2: 性別對數學成績的影響
def visualize_gender_differences(df):
    # 將性別編碼映射為男/女
    gender_mapping = {1: '男', 2: '女'}
    df['性別標籤'] = df['性別'].map(gender_mapping)
    
    plt.figure(figsize=(10, 6))
    
    # 按年級和性別分組計算平均得分率
    gender_grade_scores = df.groupby(['年級', '性別標籤'])['總得分率'].mean().reset_index()
    
    # 繪製分組條形圖
    sns.barplot(x='年級', y='總得分率', hue='性別標籤', data=gender_grade_scores, palette='Set1')
    plt.title('不同性別在各年級的平均數學得分率比較')
    plt.xlabel('年級')
    plt.ylabel('平均得分率')
    plt.ylim(0, 1)
    plt.legend(title='性別')
    
    plt.savefig('性別對數學成績的影響.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    # 使用小提琴圖顯示性別差異的分布
    plt.figure(figsize=(12, 8))
    for grade in sorted(df['年級'].unique()):
        plt.subplot(2, 3, grade-2)  # 從3年級開始，所以需要調整
        subset = df[df['年級'] == grade]
        sns.violinplot(x='性別標籤', y='總得分率', data=subset, palette='Set2')
        plt.title(f'{grade}年級')
        plt.ylim(0, 1)
        if grade > 5:
            plt.xlabel('性別')
        else:
            plt.xlabel('')
        if grade % 3 == 0:
            plt.ylabel('得分率')
        else:
            plt.ylabel('')
    
    plt.tight_layout()
    plt.savefig('各年級性別成績分布.png', dpi=300, bbox_inches='tight')
    plt.show()

# 視覺化3: 不同學校之間的數學表現差異
def visualize_school_differences(df):
    # 獲取前10個學校（按學生數量）
    school_counts = df['學校代碼'].value_counts().head(10).index
    filtered_df = df[df['學校代碼'].isin(school_counts)]
    
    # 計算每所學校在每個年級的平均得分率
    school_grade_scores = filtered_df.groupby(['學校代碼', '年級'])['總得分率'].mean().reset_index()
    
    # 繪製熱力圖
    pivot_table = school_grade_scores.pivot(index='學校代碼', columns='年級', values='總得分率')
    
    plt.figure(figsize=(12, 8))
    sns.heatmap(pivot_table, annot=True, cmap='YlGnBu', fmt='.2f', linewidths=.5)
    plt.title('不同學校在各年級的平均數學得分率熱力圖')
    plt.xlabel('年級')
    plt.ylabel('學校代碼')
    
    plt.savefig('學校間數學表現差異熱力圖.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    # 為每個年級繪製學校得分率的條形圖
    for grade in sorted(df['年級'].unique()):
        grade_data = filtered_df[filtered_df['年級'] == grade]
        
        if len(grade_data) > 0:
            # 為每所學校計算該年級的平均得分率
            school_scores = grade_data.groupby('學校代碼')['總得分率'].mean().sort_values(ascending=False).reset_index()
            
            plt.figure(figsize=(12, 6))
            sns.barplot(x='學校代碼', y='總得分率', data=school_scores, palette='viridis')
            plt.title(f'{grade}年級各學校的平均數學得分率')
            plt.xlabel('學校代碼')
            plt.ylabel('平均得分率')
            plt.ylim(0, 1)
            plt.xticks(rotation=45)
            
            # 在每個條形上顯示具體數值
            for i, row in school_scores.iterrows():
                plt.text(i, row['總得分率'] + 0.02, f'{row["總得分率"]:.2f}', 
                         ha='center', va='bottom')
            
            plt.tight_layout()
            plt.savefig(f'{grade}年級學校間數學表現差異.png', dpi=300, bbox_inches='tight')
            plt.show()

# 主函數
def main():
    # 定義資料檔案路徑
    file_paths = [
        '113年度_學力測驗_金門縣_數學3年級成績_202406.csv',
        '113年度_學力測驗_金門縣_數學4年級成績_202406.csv',
        '113年度_學力測驗_金門縣_數學5年級成績_202406.csv',
        '113年度_學力測驗_金門縣_數學6年級成績_202406.csv',
        '113年度_學力測驗_金門縣_數學7年級成績_202406.csv',
        '113年度_學力測驗_金門縣_數學8年級成績_202406.csv'
    ]
    
    # 載入並整合所有資料
    combined_data = load_all_data(file_paths)
    
    if combined_data is not None:
        print(f"成功載入資料，共 {len(combined_data)} 筆記錄")
        
        # 執行視覺化
        visualize_grade_differences(combined_data)
        visualize_gender_differences(combined_data)
        visualize_school_differences(combined_data)
    else:
        print("無法載入資料")

if __name__ == "__main__":
    main()