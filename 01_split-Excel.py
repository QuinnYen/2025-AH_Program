import tkinter as tk
from tkinter import filedialog, messagebox, ttk  # 添加ttk模組用於進度條
import pandas as pd
import os
import datetime # 匯入 datetime 模組
import sys # 匯入 sys 模組用於輸出診斷訊息
import threading  # 添加執行緒模組
import queue  # 添加佇列模組用於執行緒間通訊

# 調試級別設定：0=無輸出，1=重要訊息，2=詳細訊息
DEBUG_LEVEL = 1

def print_debug(message, level=1):
    """輸出診斷訊息到終端機
    
    參數:
        message: 要輸出的訊息
        level: 訊息重要性級別 (1=重要訊息, 2=詳細訊息)
    """
    if DEBUG_LEVEL >= level:
        print(f"[INFO] {message}" if level == 1 else f"[DEBUG] {message}")
        sys.stdout.flush()  # 確保立即輸出

def select_file():
    """開啟檔案對話框讓使用者選取檔案，並更新路徑標籤。"""
    file_path = filedialog.askopenfilename(
        title="選取 Excel 或 CSV 檔案",
        filetypes=(("Excel 檔案", "*.xlsx *.xls"), ("CSV 檔案", "*.csv"), ("所有檔案", "*.*"))
    )
    if file_path:
        file_path_var.set(file_path)
    else:
        file_path_var.set("")

def select_basic_data_file():
    """開啟檔案對話框讓使用者選取基本資料檔案，並更新路徑標籤。"""
    file_path = filedialog.askopenfilename(
        title="選取基本資料 Excel 或 CSV 檔案",
        filetypes=(("CSV 檔案", "*.csv"), ("Excel 檔案", "*.xlsx *.xls"),  ("所有檔案", "*.*"))
    )
    if file_path:
        basic_data_file_path_var.set(file_path)
    else:
        basic_data_file_path_var.set("")

def get_academic_year(semester_code):
    """根據開課學年期代碼（例如 1101）返回學年度字串（例如 110）。"""
    if pd.isna(semester_code):
        return "未知學期"
    try:
        code_str = str(int(semester_code)) # 確保是整數再轉字串，避免 .0
        if len(code_str) >= 3:
            return code_str[:3] # 取前三碼作為學年度
    except ValueError: # 處理無法轉換為整數的情況，例如空字串或非數字字元
        pass # 或者可以記錄錯誤或返回特定值
    return "未知學期"

def update_progress(status_message, progress_value=None):
    """更新進度條和狀態標籤"""
    if status_message:
        status_label.config(text=status_message)
    
    if progress_value is not None:
        progress_bar["value"] = progress_value
    
    # 更新UI
    root.update_idletasks()

def process_file_thread():
    """在背景執行緒中處理檔案"""
    try:
        file_path = file_path_var.get()
        basic_data_path = basic_data_file_path_var.get()
        
        if not file_path:
            messagebox.showerror("錯誤", "請先選取主要檔案！")
            update_progress("處理已停止", 0)
            process_button.config(state=tk.NORMAL)
            cancel_button.config(state=tk.DISABLED)
            return

        # 更新進度和狀態
        update_progress("開始讀取主檔案...", 5)

        # 讀取主檔案
        if file_path.endswith('.csv'):
            print_debug(f"開始讀取CSV檔案: {file_path}", level=1)
            update_progress("讀取CSV主檔案中...", 10)
            df = pd.read_csv(file_path)
        elif file_path.endswith(('.xlsx', '.xls')):
            print_debug(f"開始讀取Excel檔案: {file_path}", level=1)
            update_progress("讀取Excel主檔案中...", 10)
            df = pd.read_excel(file_path)
        else:
            messagebox.showerror("錯誤", "不支援的檔案格式！請選取 Excel 或 CSV 檔案。")
            update_progress("處理已停止", 0)
            process_button.config(state=tk.NORMAL)
            cancel_button.config(state=tk.DISABLED)
            return

        update_progress("檢查主檔案欄位...", 15)
        print_debug(f"主檔案欄位: {df.columns.tolist()}", level=2)

        if '開課學年期' not in df.columns:
            messagebox.showerror("錯誤", "檔案中找不到 '開課學年期' 欄位！")
            update_progress("處理已停止", 0)
            process_button.config(state=tk.NORMAL)
            cancel_button.config(state=tk.DISABLED)
            return
            
        # 檢查是否有學號欄位
        if '學號' not in df.columns:
            messagebox.showerror("錯誤", "主檔案中找不到 '學號' 欄位！")
            update_progress("處理已停止", 0)
            process_button.config(state=tk.NORMAL)
            cancel_button.config(state=tk.DISABLED)
            return
        
        # 讀取並合併基本資料檔案
        if basic_data_path:
            try:
                update_progress("開始讀取基本資料檔案...", 20)
                print_debug(f"開始讀取基本資料檔案: {basic_data_path}", level=1)
                print_debug(f"檔案副檔名檢查: 是否為CSV檔案? {basic_data_path.lower().endswith('.csv')}", level=2)
                
                # 根據檔案類型使用不同讀取方法
                if basic_data_path.lower().endswith('.csv'):
                    print_debug("識別為CSV檔案，使用read_csv讀取", level=1)
                    update_progress("讀取CSV基本資料檔案中...", 25)
                    # 嘗試不同的編碼和分隔符號
                    try:
                        # 先嘗試讀取標題行
                        with open(basic_data_path, 'r', encoding='utf-8') as f:
                            first_line = f.readline().strip()
                            print_debug(f"CSV首行內容: {first_line}", level=2)
                        
                        # 直接跳過第一行讀取
                        print_debug("嘗試跳過第一行作為標題行", level=2)
                        basic_df = pd.read_csv(basic_data_path, encoding='utf-8', skiprows=1)
                        print_debug(f"跳過第一行後的欄位: {basic_df.columns.tolist()}", level=2)
                            
                    except UnicodeDecodeError:
                        update_progress("UTF-8編碼失敗，嘗試big5編碼", 26)
                        print_debug("UTF-8編碼失敗，嘗試big5編碼", level=1)
                        # 直接跳過第一行讀取
                        basic_df = pd.read_csv(basic_data_path, encoding='big5', skiprows=1)
                        print_debug(f"跳過第一行(big5編碼)後的欄位: {basic_df.columns.tolist()}", level=2)
                    
                    # 如果沒有正確解析欄位，可能是分隔符號問題
                    if len(basic_df.columns) <= 2:  # 假設正確讀取時應該有多個欄位
                        update_progress("CSV欄位解析有問題，嘗試其他分隔符號", 27)
                        print_debug("CSV欄位解析可能有問題，嘗試其他分隔符號", level=1)
                        # 嘗試其他分隔符號
                        basic_df = pd.read_csv(basic_data_path, encoding='utf-8', sep=',', engine='python', skiprows=1)
                elif basic_data_path.lower().endswith(('.xlsx', '.xls')):
                    print_debug("識別為Excel檔案，使用read_excel讀取", level=1)
                    update_progress("讀取Excel基本資料檔案中...", 25)
                    basic_df = pd.read_excel(basic_data_path)
                else:
                    print_debug(f"無法識別的檔案類型: {basic_data_path}，嘗試作為CSV讀取", level=1)
                    update_progress(f"無法識別的檔案類型，嘗試作為CSV讀取", 25)
                    basic_df = pd.read_csv(basic_data_path, encoding='utf-8', engine='python')
                
                update_progress("檢查基本資料欄位...", 30)
                print_debug(f"基本資料檔案欄位: {basic_df.columns.tolist()}", level=2)
                
                # 檢查'學  號'欄位是否存在，考慮空白字符問題
                found_student_id_column = False
                student_id_column_name = None
                
                for col in basic_df.columns:
                    # 輸出欄位名稱及其ASCII代碼，用於診斷
                    print_debug(f"欄位: '{col}' ASCII: {[ord(c) for c in col]}", level=2)
                    
                    # 嘗試多種匹配方式
                    if (col.strip() == '學號' or 
                        '學號' in col.strip() or 
                        '學號' in col.replace(' ', '') or 
                        col.strip() == '學  號' or 
                        '學  號' in col):
                        found_student_id_column = True
                        student_id_column_name = col
                        print_debug(f"找到學號欄位: '{col}'", level=1)
                        break
                
                if not found_student_id_column:
                    error_msg = "基本資料檔案中找不到含有'學號'的欄位！"
                    print_debug(error_msg, level=1)
                    messagebox.showerror("錯誤", error_msg)
                    update_progress("處理已停止", 0)
                    process_button.config(state=tk.NORMAL)
                    cancel_button.config(state=tk.DISABLED)
                    return
                
                if '學院' not in basic_df.columns:
                    error_msg = "基本資料檔案中找不到 '學院' 欄位！"
                    print_debug(error_msg, level=1)
                    messagebox.showerror("錯誤", error_msg)
                    update_progress("處理已停止", 0)
                    process_button.config(state=tk.NORMAL)
                    cancel_button.config(state=tk.DISABLED)
                    return
                
                # 合併檔案，只保留學院欄位
                update_progress("準備合併資料...", 35)
                print_debug(f"使用欄位 '{student_id_column_name}' 進行合併", level=1)
                basic_df_selected = basic_df[[student_id_column_name, '學院']]
                
                # 輸出一些原始資料樣本
                print_debug(f"基本資料前5筆: \n{basic_df_selected.head(5)}", level=2)
                
                # 重命名欄位以便合併
                basic_df_selected = basic_df_selected.rename(columns={student_id_column_name: '學號'})
                
                print_debug(f"合併前資料筆數: 主檔案 {len(df)}, 基本資料 {len(basic_df_selected)}", level=1)
                
                # 合併前先將學號欄位轉為字串類型，並確保去除可能的浮點數小數點
                update_progress("處理學號格式...", 40)
                df['學號'] = df['學號'].astype(str).str.replace('.0', '', regex=False)
                basic_df_selected['學號'] = basic_df_selected['學號'].astype(str).str.replace('.0', '', regex=False)
                
                # 檢查重複資料
                print_debug(f"主檔案中學號為11057272的記錄數: {df[df['學號'] == '11057272'].shape[0]}", level=2)
                print_debug(f"基本資料中學號為11057272的記錄數: {basic_df_selected[basic_df_selected['學號'] == '11057272'].shape[0]}", level=2)
                
                # 檢查學號的重複情況
                update_progress("檢查重複學號...", 45)
                dup_student_ids = df['學號'].value_counts()
                dup_student_ids = dup_student_ids[dup_student_ids > 1]
                if not dup_student_ids.empty:
                    print_debug(f"主檔案中有重複學號，前5筆: \n{dup_student_ids.head(5)}", level=2)

                # 檢查基本資料中的重複
                basic_dup_ids = basic_df_selected['學號'].value_counts()
                basic_dup_ids = basic_dup_ids[basic_dup_ids > 1]
                if not basic_dup_ids.empty:
                    print_debug(f"基本資料中有重複學號，前5筆: \n{basic_dup_ids.head(5)}", level=2)
                    
                    # 處理基本資料中的重複學號
                    update_progress("處理基本資料中的重複學號...", 50)
                    print_debug("處理基本資料中的重複學號...", level=1)
                    
                    # 建立一個新的資料框架，用來存放處理後的學號與學院關係
                    processed_student_data = []
                    
                    # 檢查每個重複學號的學院是否相同
                    student_ids = basic_df_selected['學號'].unique()
                    total_ids = len(student_ids)
                    
                    for i, student_id in enumerate(student_ids):
                        if i % 100 == 0:  # 每處理100筆更新一次進度
                            update_progress(f"處理重複學號 {i+1}/{total_ids}...", 50 + (i/total_ids * 10))
                            
                        # 取得此學號的所有記錄
                        student_records = basic_df_selected[basic_df_selected['學號'] == student_id]
                        
                        if len(student_records) > 1:
                            # 有多筆記錄，檢查學院是否相同
                            unique_colleges = student_records['學院'].unique()
                            
                            if len(unique_colleges) > 1:
                                print_debug(f"處理學號 {student_id} 有多個不同學院: {unique_colleges.tolist()}", level=1)
                                
                                # 取得第一筆作為主要學院
                                main_college = student_records.iloc[0]['學院']
                                
                                # 其餘學院作為附屬學院，以逗號分隔
                                subsidiary_colleges = []
                                for i in range(1, len(student_records)):
                                    college = student_records.iloc[i]['學院']
                                    if college != main_college and college not in subsidiary_colleges:
                                        subsidiary_colleges.append(college)
                                
                                # 將處理後的資料加入列表
                                processed_student_data.append({
                                    '學號': student_id,
                                    '學院': main_college,
                                    '附屬學院': ','.join(subsidiary_colleges)
                                })
                            else:
                                # 學院都相同，只保留一筆
                                processed_student_data.append({
                                    '學號': student_id,
                                    '學院': unique_colleges[0],
                                    '附屬學院': ''
                                })
                        else:
                            # 只有一筆記錄
                            processed_student_data.append({
                                '學號': student_id,
                                '學院': student_records.iloc[0]['學院'],
                                '附屬學院': ''
                            })
                    
                    # 將處理後的資料轉換為DataFrame
                    basic_df_selected = pd.DataFrame(processed_student_data)
                    print_debug(f"處理後基本資料筆數: {len(basic_df_selected)}", level=1)
                    print_debug(f"處理後基本資料範例:\n{basic_df_selected.head()}", level=2)
                    
                # 合併資料前，確保沒有重複的學號
                update_progress("進行最終檢查...", 60)
                if len(basic_df_selected) != len(basic_df_selected['學號'].unique()):
                    print_debug("警告：處理後的基本資料仍有重複學號", level=1)
                
                # 檢查主檔案中的重複課程記錄
                if not dup_student_ids.empty:
                    print_debug("檢查主檔案中的重複課程記錄...", level=2)
                    for dup_id in dup_student_ids.index[:3]:  # 只檢查前3個重複學號
                        dup_courses = df[df['學號'] == dup_id][['開課學年期', '課程代碼', '課程名稱']].drop_duplicates()
                        dup_count = len(dup_courses)
                        print_debug(f"學號 {dup_id} 修了 {dup_count} 門不同課程", level=2)
                        if dup_count <= 5:  # 如果課程數少於5，顯示詳細課程
                            print_debug(f"課程詳情：\n{dup_courses}", level=2)
                
                # 輸出一些樣本進行對比檢查
                print_debug(f"主檔案學號前10筆: {df['學號'].head(10).tolist()}", level=2)
                print_debug(f"基本資料學號前10筆: {basic_df_selected['學號'].head(10).tolist()}", level=2)
                
                update_progress("合併資料中...", 65)
                df = pd.merge(df, basic_df_selected, on='學號', how='left')
                print_debug(f"合併後資料筆數: {len(df)}", level=1)
                
                # 檢查合併結果中學院欄位的非空值數量
                non_null_count = df['學院'].notna().sum()
                print_debug(f"合併後學院欄位非空值數量: {non_null_count} (佔比 {non_null_count/len(df)*100:.2f}%)", level=1)
                
                print_debug(f"已成功合併基本資料檔案中的學院資訊，合併成功率: {non_null_count}/{len(df)} ({non_null_count/len(df)*100:.2f}%)", level=1)
            except Exception as e:
                error_msg = f"合併基本資料時發生錯誤：\n{str(e)}"
                print_debug(f"錯誤: {error_msg}", level=1)
                print_debug(f"錯誤詳情: {type(e).__name__}", level=1)
                import traceback
                print_debug(traceback.format_exc(), level=2)
                messagebox.showerror("錯誤", error_msg)
                update_progress("處理已停止", 0)
                process_button.config(state=tk.NORMAL)
                cancel_button.config(state=tk.DISABLED)
                return
        
        # 移除姓名欄位
        update_progress("處理學年度資料...", 70)
        if '姓名' in df.columns:
            df = df.drop(columns=['姓名'])
            print_debug("已移除姓名欄位", level=1)

        df['學年度'] = df['開課學年期'].apply(get_academic_year)
        
        # 獲取檔案所在目錄
        input_file_dir = os.path.dirname(file_path)
        # 建立帶時間戳的主輸出資料夾
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        main_output_folder_name = f"處理結果_{timestamp}"
        main_output_path = os.path.join(input_file_dir, main_output_folder_name)
        os.makedirs(main_output_path, exist_ok=True)

        grouped = df.groupby('學年度')
        
        processed_years = []
        
        # 計算總學年度數量用於進度條
        total_years = len(grouped)
        update_progress(f"開始依學年度處理資料，共 {total_years} 個學年度...", 75)
        
        for i, (year_group, data) in enumerate(grouped):
            # 更新進度條
            progress_percent = 75 + (i / total_years * 20)  # 從75%到95%
            update_progress(f"處理 {year_group} 學年度資料 ({i+1}/{total_years})...", progress_percent)
            
            sheet_name_suffix = "資料"
            # 將學年度資料夾建立在時間戳資料夾內
            if year_group == "未知學期":
                academic_year_folder_name = "未知學期資料"
                file_basename = "未知學期資料.xlsx"
                sheet_name = f"未知學期{sheet_name_suffix}"
            else:
                academic_year_folder_name = f"{year_group}學年度"
                file_basename = f"{year_group}學年度課程資料.xlsx"
                sheet_name = f"{year_group}學年度{sheet_name_suffix}"
            
            # 完整的學年度資料夾路徑
            full_academic_year_folder_path = os.path.join(main_output_path, academic_year_folder_name)
            os.makedirs(full_academic_year_folder_path, exist_ok=True)
            
            # 完整的檔案儲存路徑
            output_file_path = os.path.join(full_academic_year_folder_path, file_basename)
            
            data_to_save = data.drop(columns=['學年度']) # 儲存前移除輔助的'學年度'欄
            
            update_progress(f"儲存 {year_group} 學年度資料到 Excel...", progress_percent)
            # 使用ExcelWriter並自動調整欄位寬度
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                data_to_save.to_excel(writer, index=False, sheet_name=sheet_name)
                
                # 獲取工作表
                worksheet = writer.sheets[sheet_name]
                
                # 設置每個欄位的寬度
                for idx, col in enumerate(data_to_save.columns):
                    # 計算欄位標題的寬度，中文字符需要更多空間
                    col_width = max(len(str(col)) * 1.5, 12)
                    
                    # 計算欄位內容的最大寬度
                    for row in range(min(len(data_to_save), 1000)):  # 限制檢查的行數以提高效率
                        cell_value = str(data_to_save.iloc[row, idx])
                        # 對於中文字符給予更多空間
                        has_chinese = any('\u4e00' <= char <= '\u9fff' for char in cell_value)
                        multiplier = 1.8 if has_chinese else 1.4
                        col_width = max(col_width, len(cell_value) * multiplier)
                    
                    # 增加額外的緩衝空間並設置欄位寬度 (最大120)
                    col_width += 2  # 增加固定緩衝空間
                    col_letter = worksheet.cell(row=1, column=idx+1).column_letter
                    worksheet.column_dimensions[col_letter].width = min(col_width, 120)
            processed_years.append(year_group)

        update_progress("完成！", 100)
        if processed_years:
            print_debug(f"檔案處理完成！已成功處理學年度：{', '.join(processed_years)}\n資料已存至資料夾：{main_output_path}", level=1)
        else:
            print_debug("沒有可處理的資料或學年度。", level=1)

    except FileNotFoundError:
        messagebox.showerror("錯誤", f"找不到檔案：{file_path}")
        update_progress("處理已停止", 0)
    except pd.errors.EmptyDataError:
        messagebox.showerror("錯誤", "檔案是空的！")
        update_progress("處理已停止", 0)
    except Exception as e:
        messagebox.showerror("錯誤", f"處理過程中發生錯誤：\n{str(e)}")
        update_progress("處理已停止", 0)
    finally:
        # 恢復按鈕狀態
        process_button.config(state=tk.NORMAL)
        cancel_button.config(state=tk.DISABLED)

def process_file():
    """開始處理檔案並顯示進度"""
    # 停用處理按鈕防止重複點擊，啟用取消按鈕
    process_button.config(state=tk.DISABLED)
    cancel_button.config(state=tk.NORMAL)
    
    # 重置進度條
    progress_bar["value"] = 0
    update_progress("準備處理檔案...", 0)
    
    # 創建並啟動新線程
    processing_thread = threading.Thread(target=process_file_thread)
    processing_thread.daemon = True  # 將線程設為守護線程，主程式結束時線程也會結束
    processing_thread.start()

def cancel_processing():
    """取消處理（目前僅恢復UI狀態）"""
    update_progress("已取消處理", 0)
    process_button.config(state=tk.NORMAL)
    cancel_button.config(state=tk.DISABLED)
    # 注意：目前無法實際停止正在進行的處理，只能更新UI狀態
    # 未來可以實現通過某種機制（如共享變數）來通知線程停止處理


# --- GUI 設定 ---
root = tk.Tk()
window_width = 600
window_height = 400  # 增加高度以容納進度條
# 將視窗置中於螢幕
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
root.title("學籍資料切割工具")
root.resizable(False, False) # 禁止調整視窗大小
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

# 檔案路徑變數
file_path_var = tk.StringVar()
basic_data_file_path_var = tk.StringVar()

# 框架
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(expand=True, fill=tk.BOTH)

# 步驟1: 選取檔案
step1_frame = tk.LabelFrame(main_frame, text="步驟 1: 選取檔案", padx=10, pady=10)
step1_frame.pack(pady=10, fill=tk.X)

select_button = tk.Button(step1_frame, text="選取主要檔案", command=select_file)
select_button.pack(side=tk.LEFT, padx=(0, 10))

file_label = tk.Label(step1_frame, textvariable=file_path_var, relief=tk.SUNKEN, anchor=tk.W, width=40)
file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

# 步驟2: 選取基本資料檔案
step2_frame = tk.LabelFrame(main_frame, text="步驟 2: 選取基本資料", padx=10, pady=10)
step2_frame.pack(pady=10, fill=tk.X)

select_basic_data_button = tk.Button(step2_frame, text="選取基本資料檔案", command=select_basic_data_file)
select_basic_data_button.pack(side=tk.LEFT, padx=(0, 10))

basic_data_file_label = tk.Label(step2_frame, textvariable=basic_data_file_path_var, relief=tk.SUNKEN, anchor=tk.W, width=40)
basic_data_file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

# 步驟3: 開始處理 (按鈕)
step3_frame = tk.LabelFrame(main_frame, text="步驟 3: 處理檔案", padx=10, pady=10)
step3_frame.pack(pady=10, fill=tk.X)

button_frame = tk.Frame(step3_frame)
button_frame.pack(fill=tk.X, pady=5)

process_button = tk.Button(button_frame, text="開始處理", command=process_file, width=15)
process_button.pack(side=tk.LEFT, padx=5)

cancel_button = tk.Button(button_frame, text="取消處理", command=cancel_processing, width=15, state=tk.DISABLED)
cancel_button.pack(side=tk.LEFT, padx=5)

# 進度顯示區
progress_frame = tk.Frame(main_frame)
progress_frame.pack(fill=tk.X, pady=10)

# 狀態標籤
status_label = tk.Label(progress_frame, text="就緒", anchor=tk.W)
status_label.pack(fill=tk.X)

# 進度條
progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate")
progress_bar.pack(fill=tk.X, pady=5)

# 程式說明標籤
info_label = tk.Label(main_frame, text="處理大量資料時請耐心等待，進度條會顯示目前處理進度", fg="blue")
info_label.pack(pady=5)

root.mainloop()
