import tkinter as tk
from tkinter import filedialog, messagebox, ttk  # 添加ttk用於進度條
import pandas as pd
import os
import numpy as np
import threading  # 添加執行緒模組

class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.window_width = 500
        self.window_height = 400  # 增加高度以容納進度條
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.x_position = (self.screen_width - self.window_width) // 2
        self.y_position = (self.screen_height - self.window_height) // 2
        self.root.resizable(False, False) # 禁止調整視窗大小
        self.root.title("Excel 資料處理程式")
        self.root.geometry(f"{self.window_width}x{self.window_height}+{self.x_position}+{self.y_position}")
        self.excel_path = ""
        self.processing = False  # 追蹤是否正在處理中
        
        # 創建主框架
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        # 標題
        title_label = tk.Label(self.main_frame, text="Excel 資料處理", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 步驟 1: 匯入 Excel
        step1_frame = tk.LabelFrame(self.main_frame, text="步驟 1: 匯入 Excel 檔案", font=("Arial", 12))
        step1_frame.pack(fill=tk.X, pady=10)
        
        self.file_path_label = tk.Label(step1_frame, text="尚未選擇檔案", width=40, anchor="w")
        self.file_path_label.pack(side=tk.LEFT, padx=10, pady=10)
        
        import_button = tk.Button(step1_frame, text="選擇檔案", command=self.import_excel)
        import_button.pack(side=tk.RIGHT, padx=10, pady=10)
        
        # 步驟 2: 執行處理
        step2_frame = tk.LabelFrame(self.main_frame, text="步驟 2: 執行處理", font=("Arial", 12))
        step2_frame.pack(fill=tk.X, pady=10)
        
        # 按鈕框架，包含執行和取消按鈕
        button_frame = tk.Frame(step2_frame)
        button_frame.pack(padx=10, pady=10)
        
        self.process_button = tk.Button(button_frame, text="執行", command=self.process_excel, width=15)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = tk.Button(button_frame, text="取消", command=self.cancel_processing, width=15, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        
        # 進度指示區域
        progress_frame = tk.Frame(self.main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        # 狀態顯示
        self.status_label = tk.Label(progress_frame, text="就緒", font=("Arial", 10), anchor="w")
        self.status_label.pack(fill=tk.X)
        
        # 進度條
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=100, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # 提示標籤
        info_label = tk.Label(self.main_frame, text="處理大量資料時請耐心等待，進度條會顯示目前處理進度", fg="blue")
        info_label.pack(pady=5)
    
    def update_progress(self, status_message, progress_value=None):
        """更新進度條和狀態標籤"""
        if status_message:
            self.status_label.config(text=status_message)
        
        if progress_value is not None:
            self.progress_bar["value"] = progress_value
        
        # 更新UI
        self.root.update_idletasks()
    
    def import_excel(self):
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx;*.xls")]
        )
        
        if file_path:
            self.excel_path = file_path
            self.file_path_label.config(text=os.path.basename(file_path))
            self.status_label.config(text="已選擇檔案: " + os.path.basename(file_path))
            self.update_progress("已選擇檔案: " + os.path.basename(file_path), 0)
    
    def cancel_processing(self):
        """取消處理（目前僅恢復UI狀態）"""
        self.update_progress("已取消處理", 0)
        self.process_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.processing = False
        # 注意：目前無法實際停止正在進行的處理，只能更新UI狀態
    
    def process_excel(self):
        """開始處理Excel檔案，在背景執行緒中執行"""
        if not self.excel_path:
            messagebox.showerror("錯誤", "請先選擇 Excel 檔案")
            print("錯誤: 未選擇Excel檔案")  # 在終端機列印錯誤
            return
        
        # 停用執行按鈕，啟用取消按鈕
        self.process_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        
        # 重置進度條
        self.progress_bar["value"] = 0
        self.update_progress("準備處理檔案...", 0)
        
        # 設置處理中狀態
        self.processing = True
        
        # 創建並啟動背景執行緒
        processing_thread = threading.Thread(target=self.process_excel_thread)
        processing_thread.daemon = True  # 設為守護執行緒，主程式結束時執行緒也會結束
        processing_thread.start()
    
    def process_excel_thread(self):
        """在背景執行緒中處理Excel檔案"""
        try:
            # 讀取原始 Excel 檔案
            self.update_progress("正在讀取Excel檔案...", 10)
            df = pd.read_excel(self.excel_path)
            print(f"成功讀取Excel檔案，共有 {len(df)} 筆資料")
            print(f"檔案包含欄位: {df.columns.tolist()}")  # 列印所有欄位名稱進行調試
            
            # 檢查是否有學號欄位，這是必須的
            self.update_progress("檢查必要欄位...", 15)
            if "學號" not in df.columns:
                raise ValueError("Excel檔案中找不到'學號'欄位，請確認資料格式")
            
            # 先提取每個學號的基本資料（只保留每個學號的第一筆資料）
            self.update_progress("提取基本資料...", 25)
            student_info = df.drop_duplicates(subset=["學號"]).copy()
            print(f"去重後剩下 {len(student_info)} 位學生")
            
            # 建立必要的欄位並確保它們存在
            self.update_progress("確認資料欄位結構...", 30)
            needed_columns = ["學號"]
            
            # 處理學院欄位
            if "學院" in df.columns:
                needed_columns.append("學院")
            else:
                # 如果不存在學院欄位，添加空白學院欄位
                student_info["學院"] = ""
                print("警告: 找不到'學院'欄位，將使用空值")
                
            # 處理科系欄位
            self.update_progress("處理科系資訊...", 35)
            if "學生系級" in df.columns:
                needed_columns.append("學生系級")
                # 只選取需要的欄位
                student_info = student_info[needed_columns].copy()
                # 將「學生系級」改名為「科系」
                student_info = student_info.rename(columns={"學生系級": "科系"})
            else:
                # 只選取需要的欄位
                student_info = student_info[needed_columns].copy()
                # 如果不存在科系欄位，添加空白科系欄位
                student_info["科系"] = ""
                print("警告: 找不到'學生系級'欄位，將使用空值")
            
            # 如果有學院欄位，确保它出現在最前面
            if "學院" in student_info.columns:
                # 重新排列欄位，學院放第一位
                cols = ["學院"]
                for col in student_info.columns:
                    if col != "學院":
                        cols.append(col)
                student_info = student_info[cols]
            
            # 建立結果DataFrame，添加空白欄位
            self.update_progress("建立結果資料框架...", 40)
            result_df = student_info.copy()
            result_df["一般必修"] = ""
            result_df["一般選修"] = ""
            result_df["通識必修"] = ""
            result_df["通識選修"] = ""
            
            # 確認是否有課程代碼和成績相關欄位
            self.update_progress("處理通識課程資料...", 50)
            if "課程代碼" in df.columns and "成績" in df.columns:
                # 處理通識選修 (GE)
                self.update_progress("處理通識選修課程...", 60)
                ge_df = df[df["課程代碼"].str.startswith("GE", na=False)]
                if not ge_df.empty:
                    print(f"找到 {len(ge_df)} 筆通識選修課程記錄")
                    # 計算每個學生的通識選修平均成績，保留到小數點後兩位
                    ge_avg = ge_df.groupby("學號")["成績"].mean().round(2).reset_index()
                    ge_avg = ge_avg.rename(columns={"成績": "通識選修"})
                    # 合併到結果DataFrame
                    result_df = pd.merge(result_df, ge_avg, on="學號", how="left")
                    # 用合併的成績列替換原始的空白列
                    if "通識選修_y" in result_df.columns:  # 檢查合併後的欄位是否存在
                        result_df["通識選修"] = result_df["通識選修_y"].fillna("")
                        # 刪除多餘的列
                        result_df = result_df.drop(columns=["通識選修_y"])
                        if "通識選修_x" in result_df.columns:
                            result_df = result_df.drop(columns=["通識選修_x"])
                
                # 處理通識必修 (GQ)
                self.update_progress("處理通識必修課程...", 70)
                
                # 定義特定通識必修課程列表
                specific_gq_courses = [
                    "自然科學與人工智慧",
                    "運算思維與程式設計",
                    "文學經典閱讀",
                    "語文與修辭"
                ]
                
                # 檢查是否有課程名稱欄位
                if "課程名稱" not in df.columns:
                    print("警告: 找不到'課程名稱'欄位，無法依據課程名稱識別特定通識必修課程")
                    gq_df = df[df["課程代碼"].str.startswith("GQ", na=False)]
                else:
                    # 同時包含課程代碼以GQ開頭和特定名稱的課程
                    gq_by_code = df[df["課程代碼"].str.startswith("GQ", na=False)]
                    gq_by_name = df[df["課程名稱"].isin(specific_gq_courses)]
                    # 合併兩個DataFrame，並去除重複項
                    gq_df = pd.concat([gq_by_code, gq_by_name]).drop_duplicates()
                    print(f"透過課程代碼找到 {len(gq_by_code)} 筆，透過課程名稱找到 {len(gq_by_name)} 筆通識必修課程記錄")
                
                if not gq_df.empty:
                    print(f"總共找到 {len(gq_df)} 筆通識必修課程記錄")
                    # 計算每個學生的通識必修平均成績，保留到小數點後兩位
                    gq_avg = gq_df.groupby("學號")["成績"].mean().round(2).reset_index()
                    gq_avg = gq_avg.rename(columns={"成績": "通識必修"})
                    # 合併到結果DataFrame
                    result_df = pd.merge(result_df, gq_avg, on="學號", how="left")
                    # 用合併的成績列替換原始的空白列
                    if "通識必修_y" in result_df.columns:  # 檢查合併後的欄位是否存在
                        result_df["通識必修"] = result_df["通識必修_y"].fillna("")
                        # 刪除多餘的列
                        result_df = result_df.drop(columns=["通識必修_y"])
                        if "通識必修_x" in result_df.columns:
                            result_df = result_df.drop(columns=["通識必修_x"])
            else:
                print("警告: 找不到'課程代碼'或'成績'欄位，無法處理通識課程")
            
            # 處理一般必修和一般選修課程
            self.update_progress("處理一般必修和一般選修課程...", 75)
            
            # 確認是否有必要的欄位
            if "課程代碼" in df.columns and "成績" in df.columns and "必選修" in df.columns:
                # 首先，將所有通識課程的索引集合起來，這些課程將被排除
                # 通識選修 (GE)
                ge_indices = df[df["課程代碼"].str.startswith("GE", na=False)].index
                
                # 通識必修 (GQ)
                if "課程名稱" not in df.columns:
                    gq_indices = df[df["課程代碼"].str.startswith("GQ", na=False)].index
                else:
                    gq_by_code_indices = df[df["課程代碼"].str.startswith("GQ", na=False)].index
                    gq_by_name_indices = df[df["課程名稱"].isin(specific_gq_courses)].index
                    gq_indices = gq_by_code_indices.union(gq_by_name_indices)
                
                # 所有要排除的通識課程索引
                exclude_indices = ge_indices.union(gq_indices)
                
                # 取得排除通識課程後的DataFrame
                non_ge_gq_df = df.drop(exclude_indices).copy()
                print(f"排除通識課程後剩餘 {len(non_ge_gq_df)} 筆課程記錄")
                
                # 處理一般必修課程
                required_df = non_ge_gq_df[
                    (non_ge_gq_df["必選修"].str.contains("必修", na=False)) | 
                    (non_ge_gq_df["必選修"].str.contains("教必", na=False))
                ]
                
                if not required_df.empty:
                    print(f"找到 {len(required_df)} 筆一般必修課程記錄")
                    # 計算每個學生的一般必修平均成績，保留到小數點後兩位
                    required_avg = required_df.groupby("學號")["成績"].mean().round(2).reset_index()
                    required_avg = required_avg.rename(columns={"成績": "一般必修"})
                    # 合併到結果DataFrame
                    result_df = pd.merge(result_df, required_avg, on="學號", how="left")
                    # 用合併的成績列替換原始的空白列
                    if "一般必修_y" in result_df.columns:
                        result_df["一般必修"] = result_df["一般必修_y"].fillna("")
                        # 刪除多餘的列
                        result_df = result_df.drop(columns=["一般必修_y"])
                        if "一般必修_x" in result_df.columns:
                            result_df = result_df.drop(columns=["一般必修_x"])
                
                # 處理一般選修課程
                elective_df = non_ge_gq_df[
                    (non_ge_gq_df["必選修"].str.contains("選修", na=False)) | 
                    (non_ge_gq_df["必選修"].str.contains("教選", na=False))
                ]
                
                if not elective_df.empty:
                    print(f"找到 {len(elective_df)} 筆一般選修課程記錄")
                    # 計算每個學生的一般選修平均成績，保留到小數點後兩位
                    elective_avg = elective_df.groupby("學號")["成績"].mean().round(2).reset_index()
                    elective_avg = elective_avg.rename(columns={"成績": "一般選修"})
                    # 合併到結果DataFrame
                    result_df = pd.merge(result_df, elective_avg, on="學號", how="left")
                    # 用合併的成績列替換原始的空白列
                    if "一般選修_y" in result_df.columns:
                        result_df["一般選修"] = result_df["一般選修_y"].fillna("")
                        # 刪除多餘的列
                        result_df = result_df.drop(columns=["一般選修_y"])
                        if "一般選修_x" in result_df.columns:
                            result_df = result_df.drop(columns=["一般選修_x"])
            else:
                print("警告: 找不到'課程代碼'、'成績'或'必選修'欄位，無法處理一般必修和一般選修課程")
            
            # 排序欄位順序為：學院、科系、學號、一般必修、一般選修、通識必修、通識選修
            self.update_progress("調整欄位順序...", 85)
            
            # 確定要輸出的欄位順序
            output_columns = []
            if "學院" in result_df.columns:
                output_columns.append("學院")
            output_columns.extend(["科系", "學號"])
            
            # 加入成績欄位
            output_columns.extend(["一般必修", "一般選修", "通識必修", "通識選修"])
            
            # 只保留指定欄位並按照指定順序排列
            result_df = result_df[output_columns]
            
            # 如果有學院欄位，按學院排序
            self.update_progress("排序資料...", 90)
            if "學院" in result_df.columns:
                result_df = result_df.sort_values(by=["學院", "科系", "學號"])
                print("已按學院、科系和學號排序")
            else:
                result_df = result_df.sort_values(by=["科系", "學號"])
                print("已按科系和學號排序")
            
            # 儲存新的 Excel 檔案
            self.update_progress("正在儲存結果...", 95)
            output_dir = os.path.dirname(self.excel_path)
            output_filename = os.path.splitext(os.path.basename(self.excel_path))[0] + "_處理結果.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            
            # 使用ExcelWriter並自動調整欄位寬度
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='處理結果')
                
                # 獲取工作表
                worksheet = writer.sheets['處理結果']
                
                # 設置每個欄位的寬度
                for idx, col in enumerate(result_df.columns):
                    # 計算欄位標題的寬度 (中文字符需要更寬的空間)
                    col_width = len(str(col)) * 3  # 從2增加到3，使標題更寬
                    
                    # 計算欄位內容的最大寬度
                    for row in range(len(result_df)):
                        cell_value = str(result_df.iloc[row, idx])
                        # 計算內容寬度 (對中文字符使用2.5倍寬度，對數字使用較大的寬度)
                        content_width = 0
                        for char in cell_value:
                            if '\u4e00' <= char <= '\u9fff':  # 檢查是否為中文字符
                                content_width += 2.5  # 中文字符從2.0增加到2.5倍寬度
                            elif char.isdigit() or char == '.':
                                content_width += 1.2  # 數字從0.8增加到1.2倍寬度
                            else:
                                content_width += 1.5  # 其他字符從1.2增加到1.5倍寬度
                        
                        col_width = max(col_width, content_width)
                    
                    # 設置欄位寬度 (提高最小值和最大值)
                    min_width = 12  # 最小寬度從8增加到12
                    max_width = 60  # 最大寬度從50增加到60
                    col_width = max(min_width, min(col_width, max_width))
                    
                    # 應用列寬
                    col_letter = worksheet.cell(row=1, column=idx+1).column_letter
                    worksheet.column_dimensions[col_letter].width = col_width
            
            self.update_progress(f"處理完成! 已儲存至: {output_filename}", 100)
            messagebox.showinfo("成功", f"資料處理完成!\n已儲存至: {output_path}")
            print(f"成功: 已處理Excel檔案並儲存至 {output_path}")  # 在終端機列印成功訊息
            
        except Exception as e:
            error_message = f"處理過程中發生錯誤: {str(e)}"
            print(f"錯誤: {error_message}")  # 將錯誤訊息列印在終端機
            messagebox.showerror("錯誤", f"處理過程中發生錯誤:\n{str(e)}")
            self.update_progress("處理失敗", 0)
        finally:
            # 無論成功或失敗，都恢復按鈕狀態
            self.process_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            self.processing = False

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()
