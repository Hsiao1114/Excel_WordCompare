import pandas as pd
import difflib

# 保持 find_string_differences 函式不變
def find_string_differences(s1, s2):
    """
    比較兩個字串並找出只存在於其中一個字串中的不同字元。
    ... (函式內容同上，此處省略) ...
    """
    str1 = str(s1) if pd.notna(s1) else ""
    str2 = str(s2) if pd.notna(s2) else ""

    matcher = difflib.SequenceMatcher(None, str1, str2)
    
    diff_chars = []
    
    for opcode, a_start, a_end, b_start, b_end in matcher.get_opcodes():
        if opcode == 'replace':
            diff_in_s1 = str1[a_start:a_end]
            diff_in_s2 = str2[b_start:b_end]
            diff_chars.extend(list(diff_in_s1))
            diff_chars.extend(list(diff_in_s2))
            
        elif opcode == 'delete':
            diff_in_s1 = str1[a_start:a_end]
            diff_chars.extend(list(diff_in_s1))
            
        elif opcode == 'insert':
            diff_in_s2 = str2[b_start:b_end]
            diff_chars.extend(list(diff_in_s2))

    return ", ".join(sorted(list(set(diff_chars))))


def compare_excel_columns():
    """主程式：處理 Excel 檔案比較。"""
    
    print("--- Excel 兩行字串差異比較程式 ---")
    
    # 1. 程式剛開頭指定參數
    # 在讀取輸入時，使用 strip() 移除首尾多餘的空格
    file_path = input("請輸入 Excel 檔案路徑 (例如: data.xlsx): ").strip()
    sheet_name_input = input("請輸入工作表名稱 (例如: Sheet1): ").strip()
    col1_name_input = input("請輸入要比較的**第一行**名稱 (例如: 歌詞-原始): ").strip() # 提醒用戶輸入標題而非 A/B
    col2_name_input = input("請輸入要比較的**第二行**名稱 (例如: 歌詞-修改): ").strip()
    
    output_col_name = "差異字元"
    output_file_name = "差異比對結果.xlsx"

    try:
        # 2. 讀取 Excel 檔案
        # header=0 表示第一行是標題 (這是預設值，但明確寫出有助於理解)
        df = pd.read_excel(file_path, sheet_name=sheet_name_input, header=0)
        
        # 確保欄位名稱不包含任何多餘空格（這是一個常見的 Excel 匯入問題）
        df.columns = df.columns.str.strip() 

    except FileNotFoundError:
        print(f"\n錯誤: 找不到檔案 '{file_path}'。請檢查路徑是否正確。")
        return
    except ValueError as e:
        # 如果是工作表名稱錯誤，ValueError 會包含相關訊息
        if "No sheet named" in str(e):
            print(f"\n錯誤: 找不到工作表 '{sheet_name_input}'。請檢查工作表名稱是否正確 (已自動去除空格)。")
        else:
            print(f"\n讀取檔案時發生錯誤: {e}")
        return
    except Exception as e:
        print(f"\n讀取檔案時發生錯誤: {e}")
        return

    # 3. 檢查欄位名稱是否存在
    if col1_name_input in df.columns and col2_name_input in df.columns:
        print(f"\n正在比較 '{col1_name_input}' 和 '{col2_name_input}' 兩欄位的字串差異...")
        
        # 執行比較邏輯
        df[output_col_name] = df.apply(
            lambda row: find_string_differences(row[col1_name_input], row[col2_name_input]), 
            axis=1
        )
        
        # 針對兩欄字串完全相同或兩欄皆為空值的情況，將結果設為空白
        df.loc[(df[col1_name_input].astype(str) == df[col2_name_input].astype(str)) | 
               (df[output_col_name] == ''), output_col_name] = ""
        
        print("比較完成。差異字元已寫入新的欄位。")
        
        # 4. 寫入新的 Excel 檔案
        try:
            df.to_excel(output_file_name, index=False)
            print(f"\n結果已成功寫入檔案: '{output_file_name}'")
        except Exception as e:
            print(f"\n寫入檔案時發生錯誤: {e}")
            
    else:
        # **新增的除錯功能**：如果找不到欄位，列出所有可用欄位
        print("\n=======================================================")
        print("!! 錯誤: 找不到您指定的欄位名稱。")
        print(f"您輸入的第一行名稱是: '{col1_name_input}'")
        print(f"您輸入的第二行名稱是: '{col2_name_input}'")
        print("\n請參考以下**所有有效的欄位標題**，並重新執行程式：")
        for col in df.columns:
            print(f"-> '{col}'")
        print("=======================================================")
        
if __name__ == "__main__":
    compare_excel_columns()