#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T-test分析GUI介面
用於分析課程成績資料的統計比較
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
from scipy import stats
import os
import logging
import datetime
import sys
import traceback

# 檢查並處理Excel支援
try:
    import openpyxl
except ImportError:
    openpyxl = None

# 設定日誌系統
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(f'ttest_debug_{datetime.datetime.now().strftime("%Y%m%d")}.log', encoding='utf-8')
    ]
)

logger = logging.getLogger(__name__)

class TTestAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("T-test分析工具")
        
        logger.info("初始化T-test分析工具")
        
        # 設定視窗大小並置中
        window_width = 1200
        window_height = 800
        
        # 取得螢幕大小
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        # 計算置中位置
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        
        # 設定視窗位置和大小
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        logger.debug(f"視窗大小設定為 {window_width}x{window_height}，位置 ({center_x}, {center_y})")
        
        # 資料變數
        self.data = None
        self.results = {}
        self.current_analysis_result = None  # 儲存當前分析結果
        self.progress_window = None
        self.progress_var = None
        self.progress_label_var = None
        
        # 建立主要介面
        self.create_widgets()
        logger.info("GUI介面初始化完成")
    
    def create_progress_window(self, title="執行中...", total_steps=100):
        """建立進度視窗"""
        logger.debug(f"建立進度視窗: {title}, 總步驟: {total_steps}")
        
        if self.progress_window:
            self.progress_window.destroy()
        
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title(title)
        self.progress_window.geometry("400x150")
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()
        self.progress_window.resizable(False, False)
        
        # 將進度視窗置中
        self.progress_window.update_idletasks()
        x = (self.progress_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.progress_window.winfo_screenheight() // 2) - (150 // 2)
        self.progress_window.geometry(f"400x150+{x}+{y}")
        
        # 進度標籤
        self.progress_label_var = tk.StringVar()
        self.progress_label_var.set("準備開始...")
        progress_label = ttk.Label(self.progress_window, textvariable=self.progress_label_var, 
                                  font=('Arial', 10))
        progress_label.pack(pady=10)
        
        # 進度條
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(self.progress_window, variable=self.progress_var, 
                                      maximum=total_steps, length=350)
        progress_bar.pack(pady=10, padx=25)
        
        # 百分比標籤
        self.progress_percent_var = tk.StringVar()
        self.progress_percent_var.set("0%")
        percent_label = ttk.Label(self.progress_window, textvariable=self.progress_percent_var)
        percent_label.pack(pady=5)
        
        # 取消按鈕
        cancel_button = ttk.Button(self.progress_window, text="取消", 
                                  command=self.cancel_operation)
        cancel_button.pack(pady=5)
        
        self.operation_cancelled = False
        self.total_steps = total_steps
        self.current_step = 0
        
        self.progress_window.update()
    
    def update_progress(self, step, message=""):
        """更新進度"""
        if self.progress_window and not self.operation_cancelled:
            self.current_step = step
            self.progress_var.set(step)
            
            percent = int((step / self.total_steps) * 100)
            self.progress_percent_var.set(f"{percent}%")
            
            if message:
                self.progress_label_var.set(message)
                logger.debug(f"進度更新: {percent}% - {message}")
            
            self.progress_window.update()
    
    def close_progress_window(self):
        """關閉進度視窗"""
        if self.progress_window:
            logger.debug("關閉進度視窗")
            self.progress_window.destroy()
            self.progress_window = None
    
    def cancel_operation(self):
        """取消操作"""
        self.operation_cancelled = True
        logger.info("用戶取消操作")
        self.close_progress_window()
        
    def create_widgets(self):
        # 建立主要框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 檔案匯入區域
        file_frame = ttk.LabelFrame(main_frame, text="檔案匯入", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(file_frame, text="選擇資料檔案", 
                  command=self.load_file).grid(row=0, column=0, padx=5)
        self.file_label = ttk.Label(file_frame, text="未選擇檔案")
        self.file_label.grid(row=0, column=1, padx=10)
        
        # 資料預覽區域
        preview_frame = ttk.LabelFrame(main_frame, text="資料預覽", padding="10")
        preview_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 建立Treeview顯示資料
        self.tree = ttk.Treeview(preview_frame, height=6)
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滾動條
        scrollbar_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        
        scrollbar_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(xscrollcommand=scrollbar_x.set)
        
        # 實驗組選擇區域
        experiment_frame = ttk.LabelFrame(main_frame, text="實驗組選擇", padding="10")
        experiment_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 建立滾動條容器
        canvas = tk.Canvas(experiment_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(experiment_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 添加滑鼠滾輪支援
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.create_experiment_widgets(scrollable_frame)
        
        # 結果顯示區域
        result_frame = ttk.LabelFrame(main_frame, text="分析結果", padding="10")
        result_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 結果文字區域
        self.result_text = scrolledtext.ScrolledText(result_frame, width=60, height=18)
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 輸出報表按鈕
        export_button_frame = ttk.Frame(result_frame)
        export_button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(export_button_frame, text="輸出完整分析報表", 
                  command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_button_frame, text="清空結果", 
                  command=self.clear_results).pack(side=tk.LEFT, padx=5)
        
        # 設定網格權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=2)
        main_frame.rowconfigure(2, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
    def create_experiment_widgets(self, parent):
        # 第一類：課程類型比較（配對t-test）
        group1_frame = ttk.LabelFrame(parent, text="第一類：課程類型比較（配對t-test）", padding="5")
        group1_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=2)
        
        # 實驗組合1：必修vs選修
        ttk.Label(group1_frame, text="實驗組合1：必修vs選修").grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        ttk.Button(group1_frame, text="一般必修vs一般選修", 
                  command=lambda: self.run_paired_ttest("一般必修", "一般選修")).grid(row=1, column=0, sticky=tk.W, pady=1, padx=(0,5))
        ttk.Button(group1_frame, text="通識必修vs通識選修", 
                  command=lambda: self.run_paired_ttest("通識必修", "通識選修")).grid(row=1, column=1, sticky=tk.W, pady=1)
        
        ttk.Button(group1_frame, text="所有必修vs所有選修", 
                  command=self.compare_all_required_vs_elective).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=1)
        
        # 實驗組合2：一般vs通識
        ttk.Label(group1_frame, text="實驗組合2：一般vs通識").grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(10,0))
        
        ttk.Button(group1_frame, text="一般必修vs通識必修", 
                  command=lambda: self.run_paired_ttest("一般必修", "通識必修")).grid(row=4, column=0, sticky=tk.W, pady=1, padx=(0,5))
        ttk.Button(group1_frame, text="一般選修vs通識選修", 
                  command=lambda: self.run_paired_ttest("一般選修", "通識選修")).grid(row=4, column=1, sticky=tk.W, pady=1)
        
        # 移除重複分析：所有一般vs所有通識（與核心專業vs博雅素養等價）
        
        # 第二類：群體間比較（獨立樣本t-test）
        group2_frame = ttk.LabelFrame(parent, text="第二類：群體間比較（獨立樣本t-test）", padding="5")
        group2_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Label(group2_frame, text="實驗組合3：學院間比較").grid(row=0, column=0, sticky=tk.W)
        
        # 學院選擇下拉選單
        college_frame = ttk.Frame(group2_frame)
        college_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(college_frame, text="學院1:").grid(row=0, column=0, sticky=tk.W)
        self.college1_var = tk.StringVar()
        self.college1_combo = ttk.Combobox(college_frame, textvariable=self.college1_var, 
                                          values=["理學院", "工學院", "商學院", "設計學院", "人文與教育學院", "法學院", "電機資訊學院"],
                                          state="readonly", width=15)
        self.college1_combo.grid(row=0, column=1, padx=5)
        self.college1_combo.set("理學院")
        
        ttk.Label(college_frame, text="學院2:").grid(row=0, column=2, sticky=tk.W, padx=(10,0))
        self.college2_var = tk.StringVar()
        self.college2_combo = ttk.Combobox(college_frame, textvariable=self.college2_var,
                                          values=["理學院", "工學院", "商學院", "設計學院", "人文與教育學院", "法學院", "電機資訊學院"],
                                          state="readonly", width=15)
        self.college2_combo.grid(row=0, column=3, padx=5)
        self.college2_combo.set("商學院")
        
        # 課程類型選擇
        course_frame = ttk.Frame(group2_frame)
        course_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(course_frame, text="課程類型:").grid(row=0, column=0, sticky=tk.W)
        self.course_type_var = tk.StringVar()
        self.course_type_combo = ttk.Combobox(course_frame, textvariable=self.course_type_var,
                                             values=["一般必修", "一般選修", "通識必修", "通識選修", "通識課程"],
                                             state="readonly", width=15)
        self.course_type_combo.grid(row=0, column=1, padx=5)
        self.course_type_combo.set("一般必修")
        
        ttk.Button(group2_frame, text="執行學院比較", 
                  command=self.compare_selected_colleges).grid(row=3, column=0, sticky=tk.W, pady=5)
        
        ttk.Label(group2_frame, text="實驗組合4：績優vs一般學生").grid(row=4, column=0, sticky=tk.W, pady=(10,0))
        ttk.Button(group2_frame, text="高GPA vs 低GPA", 
                  command=self.compare_gpa_groups).grid(row=5, column=0, sticky=tk.W, pady=1)
        ttk.Button(group2_frame, text="科系頂尖20% vs 後段20%", 
                  command=self.compare_top_bottom_students).grid(row=6, column=0, sticky=tk.W, pady=1)
        
        # 第三類：特殊情境比較
        group3_frame = ttk.LabelFrame(parent, text="第三類：特殊情境比較", padding="5")
        group3_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Label(group3_frame, text="實驗組合5：學習適應性").grid(row=0, column=0, sticky=tk.W)
        ttk.Button(group3_frame, text="必修高分生的選修表現", 
                  command=self.analyze_required_high_performers).grid(row=1, column=0, sticky=tk.W, pady=1)
        ttk.Button(group3_frame, text="選修高分生的必修表現", 
                  command=self.analyze_elective_high_performers).grid(row=2, column=0, sticky=tk.W, pady=1)
        
        # 第四類：跨領域與投入度/穩定度分析
        group4_frame = ttk.LabelFrame(parent, text="第四類：跨領域與投入度/穩定度分析", padding="5")
        group4_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=2)

        ttk.Label(group4_frame, text="實驗組合6：跨學科領域表現（獨立樣本t-test）").grid(row=0, column=0, sticky=tk.W)
        ttk.Button(group4_frame, text="理工組 vs 人文社科（通識課程）",
                  command=lambda: self.compare_stem_vs_humanities("通識課程")).grid(row=1, column=0, sticky=tk.W, pady=1)
        ttk.Button(group4_frame, text="理工組 vs 人文社科（一般選修）",
                  command=lambda: self.compare_stem_vs_humanities("一般選修")).grid(row=1, column=1, sticky=tk.W, pady=1, padx=(10,0))

        ttk.Label(group4_frame, text="實驗組合7：學習表現穩定度（配對t-test）").grid(row=2, column=0, sticky=tk.W, pady=(10,0))
        ttk.Button(group4_frame, text="個人最高分類別 vs 最低分類別",
                  command=self.analyze_stability_max_vs_min).grid(row=3, column=0, sticky=tk.W, pady=1)

        ttk.Label(group4_frame, text="實驗組合8：主修與非主修投入度（配對t-test）").grid(row=4, column=0, sticky=tk.W, pady=(10,0))
        ttk.Button(group4_frame, text="核心專業(一般) vs 博雅素養(通識)",
                  command=self.compare_major_vs_nonmajor).grid(row=5, column=0, sticky=tk.W, pady=1)

        ttk.Label(group4_frame, text="實驗組合9：頂尖與後段學生的學習差距（獨立樣本t-test）").grid(row=6, column=0, sticky=tk.W, pady=(10,0))
        ttk.Button(group4_frame, text="(各系)頂尖20% vs 後段20%的『必修-選修』差",
                  command=self.compare_gap_top_bottom_diff).grid(row=7, column=0, sticky=tk.W, pady=1)

        parent.columnconfigure(0, weight=1)
        
    def load_file(self):
        """載入CSV或Excel檔案"""
        logger.info("開始載入檔案")
        
        # 設定預設的資料夾路徑
        default_dir = "/mnt/d/Quinn_Small_House/2025_AH/全校課程與學籍1101-1131"
        if not os.path.exists(default_dir):
            default_dir = os.getcwd()
            logger.debug(f"預設資料夾不存在，使用當前目錄: {default_dir}")
        else:
            logger.debug(f"使用預設資料夾: {default_dir}")
        
        file_path = filedialog.askopenfilename(
            title="選擇資料檔案",
            initialdir=default_dir,
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            logger.info(f"選擇檔案: {file_path}")
            
            # 建立進度視窗
            self.create_progress_window("載入檔案中...", 5)
            
            try:
                self.update_progress(1, "檢查檔案格式...")
                file_extension = os.path.splitext(file_path)[1].lower()
                logger.debug(f"檔案副檔名: {file_extension}")
                
                self.update_progress(2, "載入檔案...")
                
                if file_extension in ['.xlsx', '.xls']:
                    # 載入Excel檔案
                    if openpyxl is None and file_extension == '.xlsx':
                        raise ValueError("需要安裝openpyxl套件來支援Excel檔案。請執行: pip install openpyxl")
                    
                    logger.debug("載入Excel檔案")
                    self.data = pd.read_excel(file_path)
                    
                elif file_extension == '.csv':
                    # 載入CSV檔案，嘗試不同編碼格式
                    encodings = ['utf-8-sig', 'utf-8', 'big5', 'gbk', 'cp950']
                    logger.debug("載入CSV檔案，嘗試不同編碼")
                    
                    for encoding in encodings:
                        try:
                            logger.debug(f"嘗試編碼: {encoding}")
                            self.data = pd.read_csv(file_path, encoding=encoding)
                            logger.debug(f"成功使用編碼: {encoding}")
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        raise ValueError("無法解碼CSV檔案，請檢查檔案編碼格式")
                else:
                    raise ValueError(f"不支援的檔案格式: {file_extension}")
                
                self.update_progress(3, "驗證資料格式...")
                logger.info(f"成功載入檔案，資料大小: {self.data.shape}")
                logger.debug(f"欄位名稱: {list(self.data.columns)}")
                
                # 檢查必要欄位
                required_columns = ['學院', '科系', '學號', '一般必修', '一般選修', '通識必修', '通識選修']
                missing_columns = [col for col in required_columns if col not in self.data.columns]
                if missing_columns:
                    logger.warning(f"缺少必要欄位: {missing_columns}")
                
                self.update_progress(4, "更新介面...")
                self.file_label.config(text=f"已載入: {os.path.basename(file_path)}")
                self.display_data_preview()
                
                self.update_progress(5, "完成!")
                self.close_progress_window()
                
                messagebox.showinfo("成功", f"成功載入 {len(self.data)} 筆資料")
                logger.info(f"檔案載入完成: {len(self.data)} 筆資料")
                
            except Exception as e:
                self.close_progress_window()
                error_msg = f"載入檔案時發生錯誤: {str(e)}"
                logger.error(error_msg)
                logger.error(f"錯誤詳情: {traceback.format_exc()}")
                messagebox.showerror("錯誤", error_msg)
        else:
            logger.info("用戶取消檔案選擇")
    
    def display_data_preview(self):
        """顯示資料預覽"""
        if self.data is None:
            return
            
        # 清除現有資料
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 設定欄位
        self.tree["columns"] = list(self.data.columns)
        self.tree["show"] = "headings"
        
        # 設定欄位標題
        for col in self.data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        # 插入資料（只顯示前50筆）
        for idx, row in self.data.head(50).iterrows():
            self.tree.insert("", "end", values=list(row))
    
    def run_paired_ttest(self, col1, col2):
        """執行配對t-test"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        
        try:
            # 取得有效的配對資料
            valid_data = self.data[[col1, col2]].dropna()
            
            if len(valid_data) < 2:
                messagebox.showerror("錯誤", f"有效配對資料不足（只有{len(valid_data)}筆）")
                return
            
            # 執行配對t-test
            statistic, p_value = stats.ttest_rel(valid_data[col1], valid_data[col2])
            
            # 計算描述性統計
            desc_stats = {
                f'{col1}_mean': valid_data[col1].mean(),
                f'{col1}_std': valid_data[col1].std(),
                f'{col2}_mean': valid_data[col2].mean(),
                f'{col2}_std': valid_data[col2].std(),
                'mean_diff': valid_data[col1].mean() - valid_data[col2].mean(),
                'n_pairs': len(valid_data)
            }
            
            # 顯示結果
            self.display_ttest_result(f"配對t-test: {col1} vs {col2}", 
                                    statistic, p_value, desc_stats)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")
    
    def compare_all_required_vs_elective(self):
        """比較所有必修vs所有選修"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        
        try:
            # 計算所有必修和選修的平均成績
            required_scores = []
            elective_scores = []
            
            for idx, row in self.data.iterrows():
                req_scores = [row['一般必修'], row['通識必修']]
                elec_scores = [row['一般選修'], row['通識選修']]
                
                req_valid = [s for s in req_scores if pd.notna(s)]
                elec_valid = [s for s in elec_scores if pd.notna(s)]
                
                if req_valid and elec_valid:
                    required_scores.append(np.mean(req_valid))
                    elective_scores.append(np.mean(elec_valid))
            
            if len(required_scores) < 2:
                messagebox.showerror("錯誤", "有效配對資料不足")
                return
            
            statistic, p_value = stats.ttest_rel(required_scores, elective_scores)
            
            desc_stats = {
                '所有必修_mean': np.mean(required_scores),
                '所有必修_std': np.std(required_scores),
                '所有選修_mean': np.mean(elective_scores),
                '所有選修_std': np.std(elective_scores),
                'mean_diff': np.mean(required_scores) - np.mean(elective_scores),
                'n_pairs': len(required_scores)
            }
            
            self.display_ttest_result("配對t-test: 所有必修 vs 所有選修", 
                                    statistic, p_value, desc_stats)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")
    
    # 移除重複分析：compare_all_general_vs_general_education（與核心專業vs博雅素養等價）
    
    def compare_selected_colleges(self):
        """比較選定的學院間差異"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        
        college1 = self.college1_var.get()
        college2 = self.college2_var.get()
        course_type = self.course_type_var.get()
        
        if college1 == college2:
            messagebox.showerror("錯誤", "請選擇不同的學院進行比較")
            return
        
        try:
            college1_data = []
            college2_data = []
            
            for idx, row in self.data.iterrows():
                # 直接使用Excel中的學院欄位
                college = row['學院']
                
                # 根據課程類型取得成績
                if course_type == "通識課程":
                    # 通識課程取兩個通識欄位的平均
                    scores = [row['通識必修'], row['通識選修']]
                    valid_scores = [s for s in scores if pd.notna(s)]
                    if valid_scores:
                        score = np.mean(valid_scores)
                    else:
                        continue
                else:
                    score = row[course_type]
                    if pd.isna(score):
                        continue
                
                # 分配到對應學院
                if college == college1:
                    college1_data.append(score)
                elif college == college2:
                    college2_data.append(score)
            
            if len(college1_data) < 2 or len(college2_data) < 2:
                messagebox.showerror("錯誤", f"資料不足：{college1}有{len(college1_data)}筆，{college2}有{len(college2_data)}筆")
                return
            
            # 執行獨立樣本t-test
            statistic, p_value = stats.ttest_ind(college1_data, college2_data)
            
            desc_stats = {
                f'{college1}_mean': np.mean(college1_data),
                f'{college1}_std': np.std(college1_data),
                f'{college1}_n': len(college1_data),
                f'{college2}_mean': np.mean(college2_data),
                f'{college2}_std': np.std(college2_data),
                f'{college2}_n': len(college2_data),
                'mean_diff': np.mean(college1_data) - np.mean(college2_data)
            }
            
            self.display_ttest_result(f"獨立樣本t-test: {college1} vs {college2} ({course_type})", 
                                    statistic, p_value, desc_stats)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")
    
    def compare_stem_vs_humanities(self, course_type: str):
        """跨學科領域表現：理工組 vs 人文社科組（獨立樣本t-test）
        理工組：理學院、工學院、電機資訊學院
        人文社科組：商學院、設計學院、人文與教育學院、法學院
        course_type: "通識課程" 或 "一般選修"
        """
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        try:
            stem_colleges = {"理學院", "工學院", "電機資訊學院"}
            hum_colleges = {"商學院", "設計學院", "人文與教育學院", "法學院"}

            stem_scores = []
            hum_scores = []

            for _, row in self.data.iterrows():
                college = row.get('學院')
                if pd.isna(college):
                    continue

                if course_type == "通識課程":
                    scores = [row.get('通識必修'), row.get('通識選修')]
                    valid = [s for s in scores if pd.notna(s)]
                    if not valid:
                        continue
                    score = np.mean(valid)
                else:  # 一般選修
                    score = row.get('一般選修')
                    if pd.isna(score):
                        continue

                if college in stem_colleges:
                    stem_scores.append(score)
                elif college in hum_colleges:
                    hum_scores.append(score)

            if len(stem_scores) < 2 or len(hum_scores) < 2:
                messagebox.showerror("錯誤", f"資料不足：理工組{len(stem_scores)}筆、人文社科組{len(hum_scores)}筆")
                return

            statistic, p_value = stats.ttest_ind(stem_scores, hum_scores)
            desc_stats = {
                '理工組_mean': np.mean(stem_scores),
                '理工組_std': np.std(stem_scores),
                '理工組_n': len(stem_scores),
                '人文社科組_mean': np.mean(hum_scores),
                '人文社科組_std': np.std(hum_scores),
                '人文社科組_n': len(hum_scores),
                'mean_diff': np.mean(stem_scores) - np.mean(hum_scores)
            }

            self.display_ttest_result(f"獨立樣本t-test: 理工組 vs 人文社科組（{course_type}）", statistic, p_value, desc_stats)
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")

    def analyze_stability_max_vs_min(self):
        """學習表現穩定度：個人最高分課程類別 vs 最低分課程類別（配對t-test）"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        try:
            max_scores = []
            min_scores = []

            for _, row in self.data.iterrows():
                values = [row.get('一般必修'), row.get('一般選修'), row.get('通識必修'), row.get('通識選修')]
                valid = [v for v in values if pd.notna(v)]
                if len(valid) < 2:
                    continue
                max_scores.append(np.max(valid))
                min_scores.append(np.min(valid))

            if len(max_scores) < 2:
                messagebox.showerror("錯誤", "有效配對資料不足")
                return

            statistic, p_value = stats.ttest_rel(max_scores, min_scores)
            desc_stats = {
                '最高分_mean': float(np.mean(max_scores)),
                '最高分_std': float(np.std(max_scores)),
                '最低分_mean': float(np.mean(min_scores)),
                '最低分_std': float(np.std(min_scores)),
                'mean_diff': float(np.mean(max_scores) - np.mean(min_scores)),
                'n_pairs': len(max_scores)
            }
            self.display_ttest_result("配對t-test: 個人最高分類別 vs 最低分類別", statistic, p_value, desc_stats)
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")

    def compare_major_vs_nonmajor(self):
        """主修與非主修投入度：核心專業(一般必修+一般選修) vs 博雅素養(通識必修+通識選修)（配對t-test）"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        try:
            major_scores = []
            nonmajor_scores = []

            for _, row in self.data.iterrows():
                major = [row.get('一般必修'), row.get('一般選修')]
                liberal = [row.get('通識必修'), row.get('通識選修')]
                major_valid = [s for s in major if pd.notna(s)]
                liberal_valid = [s for s in liberal if pd.notna(s)]
                if major_valid and liberal_valid:
                    major_scores.append(float(np.mean(major_valid)))
                    nonmajor_scores.append(float(np.mean(liberal_valid)))

            if len(major_scores) < 2:
                messagebox.showerror("錯誤", "有效配對資料不足")
                return

            statistic, p_value = stats.ttest_rel(major_scores, nonmajor_scores)
            desc_stats = {
                '核心專業_mean': float(np.mean(major_scores)),
                '核心專業_std': float(np.std(major_scores)),
                '博雅素養_mean': float(np.mean(nonmajor_scores)),
                '博雅素養_std': float(np.std(nonmajor_scores)),
                'mean_diff': float(np.mean(major_scores) - np.mean(nonmajor_scores)),
                'n_pairs': len(major_scores)
            }
            self.display_ttest_result("配對t-test: 核心專業(一般) vs 博雅素養(通識)", statistic, p_value, desc_stats)
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")

    def compare_gap_top_bottom_diff(self):
        """頂尖與後段學生的學習差距：比較『必修平均 - 選修平均』的差（獨立樣本t-test）
        以各科系為單位選取頂尖20%與後段20%（依科系內GPA），聚合各系後進行整體t-test。
        """
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        try:
            top_diffs = []
            bottom_diffs = []

            for dept in self.data['科系'].dropna().unique():
                dept_data = self.data[self.data['科系'] == dept].copy()
                if len(dept_data) < 10:
                    continue

                # 科系內計算GPA與『必修平均-選修平均』
                gpa_list = []  # (index, gpa)
                diff_map = {}  # index -> diff
                for idx, row in dept_data.iterrows():
                    scores = [row.get('一般必修'), row.get('一般選修'), row.get('通識必修'), row.get('通識選修')]
                    valid_scores = [s for s in scores if pd.notna(s)]
                    if len(valid_scores) < 2:
                        continue
                    gpa = float(np.mean(valid_scores))
                    # 必修平均與選修平均
                    req_valid = [s for s in [row.get('一般必修'), row.get('通識必修')] if pd.notna(s)]
                    ele_valid = [s for s in [row.get('一般選修'), row.get('通識選修')] if pd.notna(s)]
                    if not req_valid or not ele_valid:
                        continue
                    diff = float(np.mean(req_valid) - np.mean(ele_valid))
                    gpa_list.append((idx, gpa))
                    diff_map[idx] = diff

                if len(gpa_list) < 10:
                    continue

                gpa_list.sort(key=lambda x: x[1])
                n = len(gpa_list)
                bottom_20 = int(n * 0.2)
                top_20 = int(n * 0.8)
                bottom_ids = [i for i, _ in gpa_list[:bottom_20]]
                top_ids = [i for i, _ in gpa_list[top_20:]]

                top_diffs.extend([diff_map[i] for i in top_ids if i in diff_map])
                bottom_diffs.extend([diff_map[i] for i in bottom_ids if i in diff_map])

            if len(top_diffs) < 2 or len(bottom_diffs) < 2:
                messagebox.showerror("錯誤", f"資料不足：頂尖組{len(top_diffs)}筆、後段組{len(bottom_diffs)}筆")
                return

            statistic, p_value = stats.ttest_ind(top_diffs, bottom_diffs)
            desc_stats = {
                '頂尖組_mean': float(np.mean(top_diffs)),
                '頂尖組_std': float(np.std(top_diffs)),
                '頂尖組_n': len(top_diffs),
                '後段組_mean': float(np.mean(bottom_diffs)),
                '後段組_std': float(np.std(bottom_diffs)),
                '後段組_n': len(bottom_diffs),
                'mean_diff': float(np.mean(top_diffs) - np.mean(bottom_diffs))
            }
            self.display_ttest_result("獨立樣本t-test: (各系)頂尖20% vs 後段20%的『必修-選修』差", statistic, p_value, desc_stats)
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")

    def compare_gpa_groups(self):
        """比較高GPA vs 低GPA學生"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        
        try:
            # 計算每個學生的GPA（所有成績的平均）
            gpa_data = []
            for idx, row in self.data.iterrows():
                scores = [row['一般必修'], row['一般選修'], row['通識必修'], row['通識選修']]
                valid_scores = [s for s in scores if pd.notna(s)]
                if len(valid_scores) >= 2:  # 至少要有2門課的成績
                    gpa_data.append((idx, np.mean(valid_scores)))
            
            if len(gpa_data) < 10:
                messagebox.showerror("錯誤", "有效GPA資料不足")
                return
            
            # 排序並取前30%和後30%
            gpa_data.sort(key=lambda x: x[1])
            n = len(gpa_data)
            bottom_30_percent = int(n * 0.3)
            top_30_percent = int(n * 0.7)
            
            low_gpa_indices = [x[0] for x in gpa_data[:bottom_30_percent]]
            high_gpa_indices = [x[0] for x in gpa_data[top_30_percent:]]
            
            # 比較各科目類型
            subjects = ['一般必修', '一般選修', '通識必修', '通識選修']
            results = {}
            
            for subject in subjects:
                high_scores = [self.data.loc[i, subject] for i in high_gpa_indices 
                              if pd.notna(self.data.loc[i, subject])]
                low_scores = [self.data.loc[i, subject] for i in low_gpa_indices 
                             if pd.notna(self.data.loc[i, subject])]
                
                if len(high_scores) >= 2 and len(low_scores) >= 2:
                    statistic, p_value = stats.ttest_ind(high_scores, low_scores)
                    results[subject] = {
                        'statistic': statistic,
                        'p_value': p_value,
                        'high_mean': np.mean(high_scores),
                        'low_mean': np.mean(low_scores),
                        'high_n': len(high_scores),
                        'low_n': len(low_scores)
                    }
            
            # 儲存分析結果
            self.current_analysis_result = {
                'title': "高GPA學生 vs 低GPA學生比較結果",
                'type': 'gpa_comparison',
                'results': results,
                'timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 顯示結果
            result_text = "高GPA學生 vs 低GPA學生比較結果\n"
            result_text += "=" * 50 + "\n"
            
            for subject, result in results.items():
                result_text += f"\n{subject}:\n"
                result_text += f"  高GPA組: 平均={result['high_mean']:.2f}, n={result['high_n']}\n"
                result_text += f"  低GPA組: 平均={result['low_mean']:.2f}, n={result['low_n']}\n"
                result_text += f"  t統計量: {result['statistic']:.4f}\n"
                result_text += f"  p值: {result['p_value']:.4f}\n"
                result_text += f"  {'顯著' if result['p_value'] < 0.05 else '不顯著'}\n"
            
            # 清空之前的結果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, result_text)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")
    
    def compare_top_bottom_students(self):
        """比較科系頂尖20% vs 後段20%學生"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        
        try:
            # 按科系分組
            departments = self.data['科系'].unique()
            all_results = {}
            
            for dept in departments:
                dept_data = self.data[self.data['科系'] == dept].copy()
                
                if len(dept_data) < 10:  # 科系人數太少跳過
                    continue
                
                # 計算科系內GPA
                gpa_scores = []
                for idx, row in dept_data.iterrows():
                    scores = [row['一般必修'], row['一般選修'], row['通識必修'], row['通識選修']]
                    valid_scores = [s for s in scores if pd.notna(s)]
                    if len(valid_scores) >= 2:
                        gpa_scores.append((idx, np.mean(valid_scores)))
                
                if len(gpa_scores) < 10:
                    continue
                
                # 排序並取前20%和後20%
                gpa_scores.sort(key=lambda x: x[1])
                n = len(gpa_scores)
                bottom_20_percent = int(n * 0.2)
                top_20_percent = int(n * 0.8)
                
                bottom_indices = [x[0] for x in gpa_scores[:bottom_20_percent]]
                top_indices = [x[0] for x in gpa_scores[top_20_percent:]]
                
                # 比較各科目
                dept_results = {}
                subjects = ['一般必修', '一般選修', '通識必修', '通識選修']
                
                for subject in subjects:
                    top_scores = [self.data.loc[i, subject] for i in top_indices 
                                 if pd.notna(self.data.loc[i, subject])]
                    bottom_scores = [self.data.loc[i, subject] for i in bottom_indices 
                                   if pd.notna(self.data.loc[i, subject])]
                    
                    if len(top_scores) >= 2 and len(bottom_scores) >= 2:
                        statistic, p_value = stats.ttest_ind(top_scores, bottom_scores)
                        dept_results[subject] = {
                            'statistic': statistic,
                            'p_value': p_value,
                            'top_mean': np.mean(top_scores),
                            'bottom_mean': np.mean(bottom_scores),
                            'top_n': len(top_scores),
                            'bottom_n': len(bottom_scores)
                        }
                
                if dept_results:
                    all_results[dept] = dept_results
            
            # 顯示結果
            result_text = "科系頂尖20% vs 後段20%學生比較結果\n"
            result_text += "=" * 60 + "\n"
            
            for dept, dept_results in all_results.items():
                result_text += f"\n【{dept}】\n"
                for subject, result in dept_results.items():
                    result_text += f"  {subject}:\n"
                    result_text += f"    頂尖組: 平均={result['top_mean']:.2f}, n={result['top_n']}\n"
                    result_text += f"    後段組: 平均={result['bottom_mean']:.2f}, n={result['bottom_n']}\n"
                    result_text += f"    t={result['statistic']:.4f}, p={result['p_value']:.4f}\n"
            
            # 清空之前的結果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, result_text)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")
    
    def analyze_required_high_performers(self):
        """分析必修課高分學生在選修課的表現"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        
        try:
            # 計算必修課平均成績
            required_avg = []
            for idx, row in self.data.iterrows():
                req_scores = [row['一般必修'], row['通識必修']]
                valid_req = [s for s in req_scores if pd.notna(s)]
                if valid_req:
                    required_avg.append((idx, np.mean(valid_req)))
            
            if len(required_avg) < 10:
                messagebox.showerror("錯誤", "有效必修成績資料不足")
                return
            
            # 取必修課成績前30%的學生
            required_avg.sort(key=lambda x: x[1], reverse=True)
            top_30_percent = int(len(required_avg) * 0.3)
            high_required_students = [x[0] for x in required_avg[:top_30_percent]]
            
            # 分析這些學生的選修課表現
            elective_scores = []
            for idx in high_required_students:
                row = self.data.loc[idx]
                elec_scores = [row['一般選修'], row['通識選修']]
                valid_elec = [s for s in elec_scores if pd.notna(s)]
                if valid_elec:
                    elective_scores.append(np.mean(valid_elec))
            
            if len(elective_scores) < 2:
                messagebox.showerror("錯誤", "高必修分學生的選修資料不足")
                return
            
            # 比較必修高分學生的選修成績與全體學生的選修成績
            all_elective_scores = []
            for idx, row in self.data.iterrows():
                elec_scores = [row['一般選修'], row['通識選修']]
                valid_elec = [s for s in elec_scores if pd.notna(s)]
                if valid_elec:
                    all_elective_scores.append(np.mean(valid_elec))
            
            if len(all_elective_scores) < 10:
                messagebox.showerror("錯誤", "全體選修成績資料不足")
                return
            
            # 執行t-test
            statistic, p_value = stats.ttest_ind(elective_scores, all_elective_scores)
            
            result_text = "必修課高分學生在選修課的表現分析\n"
            result_text += "=" * 50 + "\n"
            result_text += f"必修高分學生數: {len(high_required_students)}\n"
            result_text += f"其中有選修成績者: {len(elective_scores)}\n"
            result_text += f"高必修分學生選修平均: {np.mean(elective_scores):.2f}\n"
            result_text += f"全體學生選修平均: {np.mean(all_elective_scores):.2f}\n"
            result_text += f"t統計量: {statistic:.4f}\n"
            result_text += f"p值: {p_value:.4f}\n"
            result_text += f"結論: 必修高分學生的選修成績{'顯著高於' if p_value < 0.05 and statistic > 0 else '與'}全體學生平均\n"
            
            # 清空之前的結果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, result_text)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")
    
    def analyze_elective_high_performers(self):
        """分析選修課高分學生在必修課的表現"""
        if self.data is None:
            messagebox.showerror("錯誤", "請先載入資料檔案")
            return
        
        try:
            # 計算選修課平均成績
            elective_avg = []
            for idx, row in self.data.iterrows():
                elec_scores = [row['一般選修'], row['通識選修']]
                valid_elec = [s for s in elec_scores if pd.notna(s)]
                if valid_elec:
                    elective_avg.append((idx, np.mean(valid_elec)))
            
            if len(elective_avg) < 10:
                messagebox.showerror("錯誤", "有效選修成績資料不足")
                return
            
            # 取選修課成績前30%的學生
            elective_avg.sort(key=lambda x: x[1], reverse=True)
            top_30_percent = int(len(elective_avg) * 0.3)
            high_elective_students = [x[0] for x in elective_avg[:top_30_percent]]
            
            # 分析這些學生的必修課表現
            required_scores = []
            for idx in high_elective_students:
                row = self.data.loc[idx]
                req_scores = [row['一般必修'], row['通識必修']]
                valid_req = [s for s in req_scores if pd.notna(s)]
                if valid_req:
                    required_scores.append(np.mean(valid_req))
            
            if len(required_scores) < 2:
                messagebox.showerror("錯誤", "高選修分學生的必修資料不足")
                return
            
            # 比較選修高分學生的必修成績與全體學生的必修成績
            all_required_scores = []
            for idx, row in self.data.iterrows():
                req_scores = [row['一般必修'], row['通識必修']]
                valid_req = [s for s in req_scores if pd.notna(s)]
                if valid_req:
                    all_required_scores.append(np.mean(valid_req))
            
            if len(all_required_scores) < 10:
                messagebox.showerror("錯誤", "全體必修成績資料不足")
                return
            
            # 執行t-test
            statistic, p_value = stats.ttest_ind(required_scores, all_required_scores)
            
            result_text = "選修課高分學生在必修課的表現分析\n"
            result_text += "=" * 50 + "\n"
            result_text += f"選修高分學生數: {len(high_elective_students)}\n"
            result_text += f"其中有必修成績者: {len(required_scores)}\n"
            result_text += f"高選修分學生必修平均: {np.mean(required_scores):.2f}\n"
            result_text += f"全體學生必修平均: {np.mean(all_required_scores):.2f}\n"
            result_text += f"t統計量: {statistic:.4f}\n"
            result_text += f"p值: {p_value:.4f}\n"
            result_text += f"結論: 選修高分學生的必修成績{'顯著高於' if p_value < 0.05 and statistic > 0 else '與'}全體學生平均\n"
            
            # 清空之前的結果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, result_text)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分析時發生錯誤: {str(e)}")
    
    def display_ttest_result(self, title, statistic, p_value, desc_stats):
        """顯示t-test結果"""
        # 清空之前的結果
        self.result_text.delete(1.0, tk.END)
        
        # 判斷顯著性
        if p_value < 0.001:
            significance = "極顯著 (p < 0.001)"
        elif p_value < 0.01:
            significance = "高度顯著 (p < 0.01)"
        elif p_value < 0.05:
            significance = "顯著 (p < 0.05)"
        else:
            significance = "不顯著 (p >= 0.05)"
        
        # 儲存當前分析結果
        self.current_analysis_result = {
            'title': title,
            'statistic': statistic,
            'p_value': p_value,
            'significance': significance,
            'desc_stats': desc_stats,
            'timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        result_text = f"{title}\n"
        result_text += "=" * 50 + "\n"
        
        # 顯示描述性統計
        for key, value in desc_stats.items():
            if isinstance(value, (int, float)):
                result_text += f"{key}: {value:.2f}\n"
            else:
                result_text += f"{key}: {value}\n"
        
        result_text += f"\nt統計量: {statistic:.4f}\n"
        result_text += f"p值: {p_value:.4f}\n"
        result_text += f"顯著性: {significance}\n"
        result_text += "\n" + "-" * 50 + "\n\n"
        
        # 將結果添加到文字區域
        self.result_text.insert(tk.END, result_text)
        self.result_text.see(tk.END)
    
    def run_all_analyses(self, progress_callback=None):
        """執行所有可能的分析並返回結果"""
        if self.data is None:
            logger.error("無資料可分析")
            return None
        
        logger.info("開始執行所有統計分析")
        all_results = {}
        current_step = 0
        
        # 計算總步驟數
        # 第一類：基礎課程類型比較（4個配對測試）
        basic_paired_tests = 4
        # 第二類：制度性分析（1個）
        institutional_analysis = 1  # 所有必修vs所有選修
        # 第三類：學科性質分析（1個）
        discipline_analysis = 1  # 專業課程vs通識課程
        # 第四類：個人學習穩定度分析（1個）
        stability_analysis = 1  # 最高vs最低分類別
        # 第五類：跨學科領域比較（1個）
        interdisciplinary_analysis = 1  # 理工vs人文社科
        # 第六類：學習能力分層分析（4個）
        performance_tier_analysis = 4  # 6a.各系頂尖vs後段各科目 + 6b.必修選修差 + 6c.高低GPA + 6d.必修高分學生選修表現
        # 第七類：學院間比較分析
        colleges_count = 7
        college_combinations = (colleges_count * (colleges_count - 1)) // 2  # 21組合
        course_types_count = 4
        college_comparison_tests = college_combinations * course_types_count  # 84個測試
        
        total_steps = (basic_paired_tests + institutional_analysis + discipline_analysis + 
                      stability_analysis + interdisciplinary_analysis + performance_tier_analysis + 
                      college_comparison_tests)
        
        logger.info(f"預計執行 {total_steps} 項分析")
        
        try:
            # ========== 第一類：基礎課程類型比較（配對t-test）==========
            # 目的：細分比較各種課程類型之間的表現差異
            paired_tests = [
                ("一般必修", "一般選修"),
                ("通識必修", "通識選修"),
                ("一般必修", "通識必修"),
                ("一般選修", "通識選修")
            ]
            
            logger.info("執行基礎課程類型配對t-test分析")
            for i, (col1, col2) in enumerate(paired_tests):
                if self.operation_cancelled:
                    logger.info("操作被取消")
                    return all_results
                
                current_step += 1
                if progress_callback:
                    progress_callback(current_step, f"配對t-test: {col1} vs {col2}")
                
                try:
                    logger.debug(f"分析 {col1} vs {col2}")
                    valid_data = self.data[[col1, col2]].dropna()
                    
                    if len(valid_data) >= 2:
                        statistic, p_value = stats.ttest_rel(valid_data[col1], valid_data[col2])
                        
                        result_key = f"配對t-test_{col1}_vs_{col2}"
                        all_results[result_key] = {
                            'type': 'paired_ttest',
                            'comparison': f"{col1} vs {col2}",
                            'statistic': statistic,
                            'p_value': p_value,
                            'mean1': valid_data[col1].mean(),
                            'std1': valid_data[col1].std(),
                            'mean2': valid_data[col2].mean(),
                            'std2': valid_data[col2].std(),
                            'mean_diff': valid_data[col1].mean() - valid_data[col2].mean(),
                            'n_pairs': len(valid_data),
                            'significance': self._get_significance(p_value)
                        }
                        logger.debug(f"完成 {col1} vs {col2}, p={p_value:.4f}")
                    else:
                        logger.warning(f"{col1} vs {col2}: 有效資料不足 ({len(valid_data)}筆)")
                        
                except Exception as e:
                    logger.error(f"分析 {col1} vs {col2} 時發生錯誤: {str(e)}")
                    continue
            
            # ========== 第二類：制度性分析（配對t-test）==========
            # 目的：從必修/選修制度角度分析整體學習成效差異
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "制度分析: 所有必修課程 vs 所有選修課程")
            
            try:
                logger.debug("分析所有必修vs所有選修")
                required_scores = []
                elective_scores = []
                
                for idx, row in self.data.iterrows():
                    req_scores = [row['一般必修'], row['通識必修']]
                    elec_scores = [row['一般選修'], row['通識選修']]
                    
                    req_valid = [s for s in req_scores if pd.notna(s)]
                    elec_valid = [s for s in elec_scores if pd.notna(s)]
                    
                    if req_valid and elec_valid:
                        required_scores.append(np.mean(req_valid))
                        elective_scores.append(np.mean(elec_valid))
                
                if len(required_scores) >= 2:
                    statistic, p_value = stats.ttest_rel(required_scores, elective_scores)
                    
                    all_results["配對t-test_所有必修_vs_所有選修"] = {
                        'type': 'paired_ttest',
                        'comparison': "所有必修 vs 所有選修（制度性分析）",
                        'statistic': statistic,
                        'p_value': p_value,
                        'mean1': np.mean(required_scores),
                        'std1': np.std(required_scores),
                        'mean2': np.mean(elective_scores),
                        'std2': np.std(elective_scores),
                        'mean_diff': np.mean(required_scores) - np.mean(elective_scores),
                        'n_pairs': len(required_scores),
                        'significance': self._get_significance(p_value)
                    }
                    logger.debug(f"完成所有必修vs所有選修, p={p_value:.4f}")
                else:
                    logger.warning(f"所有必修vs所有選修: 有效資料不足 ({len(required_scores)}筆)")
            except Exception as e:
                logger.error(f"分析所有必修vs所有選修時發生錯誤: {str(e)}")
            
            # ========== 第三類：學科性質分析（配對t-test）==========
            # 目的：比較專業教育與通識教育的整體學習成效
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "配對t-test: 專業課程整體 vs 通識課程整體")
            try:
                major_scores, liberal_scores = [], []
                for _, row in self.data.iterrows():
                    # 專業課程：一般必修+一般選修的平均
                    major = [row.get('一般必修'), row.get('一般選修')]
                    # 通識課程：通識必修+通識選修的平均
                    liberal = [row.get('通識必修'), row.get('通識選修')]
                    major_valid = [s for s in major if pd.notna(s)]
                    liberal_valid = [s for s in liberal if pd.notna(s)]
                    if major_valid and liberal_valid:
                        major_scores.append(float(np.mean(major_valid)))
                        liberal_scores.append(float(np.mean(liberal_valid)))
                if len(major_scores) >= 2:
                    statistic, p_value = stats.ttest_rel(major_scores, liberal_scores)
                    all_results["配對t-test_專業課程整體_vs_通識課程整體"] = {
                        'type': 'paired_ttest',
                        'comparison': "專業課程整體 vs 通識課程整體",
                        'statistic': statistic,
                        'p_value': p_value,
                        'mean1': float(np.mean(major_scores)),
                        'std1': float(np.std(major_scores)),
                        'mean2': float(np.mean(liberal_scores)),
                        'std2': float(np.std(liberal_scores)),
                        'mean_diff': float(np.mean(major_scores) - np.mean(liberal_scores)),
                        'n_pairs': len(major_scores),
                        'significance': self._get_significance(p_value)
                    }
            except Exception as e:
                logger.error(f"分析專業課程 vs 通識課程時發生錯誤: {str(e)}")

            # ========== 第四類：個人學習穩定度分析（配對t-test）==========
            # 目的：分析個人在不同課程類型間的學習表現穩定性
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "配對t-test: 個人最高分類別 vs 最低分類別")
            try:
                max_scores, min_scores = [], []
                for _, row in self.data.iterrows():
                    values = [row.get('一般必修'), row.get('一般選修'), row.get('通識必修'), row.get('通識選修')]
                    valid = [v for v in values if pd.notna(v)]
                    if len(valid) < 2:
                        continue
                    max_scores.append(float(np.max(valid)))
                    min_scores.append(float(np.min(valid)))
                if len(max_scores) >= 2:
                    statistic, p_value = stats.ttest_rel(max_scores, min_scores)
                    all_results["配對t-test_最高分類別_vs_最低分類別"] = {
                        'type': 'paired_ttest',
                        'comparison': "個人最高分類別 vs 最低分類別",
                        'statistic': statistic,
                        'p_value': p_value,
                        'mean1': float(np.mean(max_scores)),
                        'std1': float(np.std(max_scores)),
                        'mean2': float(np.mean(min_scores)),
                        'std2': float(np.std(min_scores)),
                        'mean_diff': float(np.mean(max_scores) - np.mean(min_scores)),
                        'n_pairs': len(max_scores),
                        'significance': self._get_significance(p_value)
                    }
            except Exception as e:
                logger.error(f"分析學習表現穩定度時發生錯誤: {str(e)}")

            # ========== 第五類：跨學科領域比較（獨立樣本t-test）==========
            # 目的：比較理工學科與人文社科學生的整體學習表現
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "跨學科領域表現: 理工組 vs 人文社科組 (整合所有課程)")
            try:
                stem_colleges = {"理學院", "工學院", "電機資訊學院"}
                hum_colleges = {"商學院", "設計學院", "人文與教育學院", "法學院"}
                stem_scores, hum_scores = [], []
                for _, row in self.data.iterrows():
                    college = row.get('學院')
                    if pd.isna(college):
                        continue
                    # 整合所有課程類型的平均分數
                    all_scores = [row.get('一般必修'), row.get('一般選修'), row.get('通識必修'), row.get('通識選修')]
                    valid_scores = [s for s in all_scores if pd.notna(s)]
                    if len(valid_scores) < 2:
                        continue
                    overall_score = float(np.mean(valid_scores))
                    if college in stem_colleges:
                        stem_scores.append(overall_score)
                    elif college in hum_colleges:
                        hum_scores.append(overall_score)
                if len(stem_scores) >= 2 and len(hum_scores) >= 2:
                    statistic, p_value = stats.ttest_ind(stem_scores, hum_scores)
                    all_results["獨立樣本t-test_理工組_vs_人文社科組_整合表現"] = {
                        'type': 'independent_ttest',
                        'comparison': "理工組 vs 人文社科組（整合所有課程表現）",
                        'statistic': statistic,
                        'p_value': p_value,
                        'mean1': float(np.mean(stem_scores)),
                        'std1': float(np.std(stem_scores)),
                        'n1': len(stem_scores),
                        'mean2': float(np.mean(hum_scores)),
                        'std2': float(np.std(hum_scores)),
                        'n2': len(hum_scores),
                        'mean_diff': float(np.mean(stem_scores) - np.mean(hum_scores)),
                        'significance': self._get_significance(p_value)
                    }
            except Exception as e:
                logger.error(f"分析跨學科領域表現時發生錯誤: {str(e)}")

            # ========== 第六類：學習能力分層分析（獨立樣本t-test）==========
            # 目的：比較頂尖學生與後段學生的學習表現差異
            
            # 6a. 各系頂尖20% vs 後段20%學生（各科目成績比較）
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "各系頂尖20% vs 後段20%學生 (各科目)")
            try:
                departments = self.data['科系'].dropna().unique()
                subjects = ['一般必修', '一般選修', '通識必修', '通識選修']
                
                for dept in departments:
                    dept_data = self.data[self.data['科系'] == dept].copy()
                    if len(dept_data) < 10:
                        continue
                    
                    # 計算科系內GPA並排序
                    gpa_scores = []
                    for idx, row in dept_data.iterrows():
                        scores = [row['一般必修'], row['一般選修'], row['通識必修'], row['通識選修']]
                        valid_scores = [s for s in scores if pd.notna(s)]
                        if len(valid_scores) >= 2:
                            gpa_scores.append((idx, np.mean(valid_scores)))
                    
                    if len(gpa_scores) < 10:
                        continue
                    
                    gpa_scores.sort(key=lambda x: x[1])
                    n = len(gpa_scores)
                    bottom_20_percent = int(n * 0.2)
                    top_20_percent = int(n * 0.8)
                    
                    bottom_indices = [x[0] for x in gpa_scores[:bottom_20_percent]]
                    top_indices = [x[0] for x in gpa_scores[top_20_percent:]]
                    
                    # 對每個科目進行 t-test
                    for subject in subjects:
                        top_scores = [self.data.loc[i, subject] for i in top_indices 
                                     if pd.notna(self.data.loc[i, subject])]
                        bottom_scores = [self.data.loc[i, subject] for i in bottom_indices 
                                       if pd.notna(self.data.loc[i, subject])]
                        
                        if len(top_scores) >= 2 and len(bottom_scores) >= 2:
                            statistic, p_value = stats.ttest_ind(top_scores, bottom_scores)
                            result_key = f"獨立樣本t-test_{dept}_頂尖vs後段_{subject}"
                            all_results[result_key] = {
                                'type': 'independent_ttest',
                                'comparison': f"{dept} 頂尖20% vs 後段20% ({subject})",
                                'statistic': statistic,
                                'p_value': p_value,
                                'mean1': float(np.mean(top_scores)),
                                'std1': float(np.std(top_scores)),
                                'n1': len(top_scores),
                                'mean2': float(np.mean(bottom_scores)),
                                'std2': float(np.std(bottom_scores)),
                                'n2': len(bottom_scores),
                                'mean_diff': float(np.mean(top_scores) - np.mean(bottom_scores)),
                                'significance': self._get_significance(p_value)
                            }
            except Exception as e:
                logger.error(f"分析各系頂尖vs後段學生時發生錯誤: {str(e)}")
            
            # 6b. 各系頂尖20% vs 後段20%的『必修-選修』差異
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "(各系)頂尖20% vs 後段20% 的『必修-選修』差")
            try:
                top_diffs, bottom_diffs = [], []
                for dept in self.data['科系'].dropna().unique():
                    dept_data = self.data[self.data['科系'] == dept].copy()
                    if len(dept_data) < 10:
                        continue
                    gpa_list = []
                    diff_map = {}
                    for idx, row in dept_data.iterrows():
                        scores = [row.get('一般必修'), row.get('一般選修'), row.get('通識必修'), row.get('通識選修')]
                        valid_scores = [s for s in scores if pd.notna(s)]
                        if len(valid_scores) < 2:
                            continue
                        gpa = float(np.mean(valid_scores))
                        req_valid = [s for s in [row.get('一般必修'), row.get('通識必修')] if pd.notna(s)]
                        ele_valid = [s for s in [row.get('一般選修'), row.get('通識選修')] if pd.notna(s)]
                        if not req_valid or not ele_valid:
                            continue
                        diff = float(np.mean(req_valid) - np.mean(ele_valid))
                        gpa_list.append((idx, gpa))
                        diff_map[idx] = diff
                    if len(gpa_list) < 10:
                        continue
                    gpa_list.sort(key=lambda x: x[1])
                    n = len(gpa_list)
                    bottom_20 = int(n * 0.2)
                    top_20 = int(n * 0.8)
                    bottom_ids = [i for i, _ in gpa_list[:bottom_20]]
                    top_ids = [i for i, _ in gpa_list[top_20:]]
                    top_diffs.extend([diff_map[i] for i in top_ids if i in diff_map])
                    bottom_diffs.extend([diff_map[i] for i in bottom_ids if i in diff_map])
                if len(top_diffs) >= 2 and len(bottom_diffs) >= 2:
                    statistic, p_value = stats.ttest_ind(top_diffs, bottom_diffs)
                    all_results["獨立樣本t-test_頂尖20%_vs_後段20%_必修減選修之差"] = {
                        'type': 'independent_ttest',
                        'comparison': "(各系)頂尖20% vs 後段20%的『必修-選修』差",
                        'statistic': statistic,
                        'p_value': p_value,
                        'mean1': float(np.mean(top_diffs)),
                        'std1': float(np.std(top_diffs)),
                        'n1': len(top_diffs),
                        'mean2': float(np.mean(bottom_diffs)),
                        'std2': float(np.std(bottom_diffs)),
                        'n2': len(bottom_diffs),
                        'mean_diff': float(np.mean(top_diffs) - np.mean(bottom_diffs)),
                        'significance': self._get_significance(p_value)
                    }
            except Exception as e:
                logger.error(f"分析頂尖與後段學生的『必修-選修』差時發生錯誤: {str(e)}")
            
            # 6c. 高GPA vs 低GPA學生比較
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "高GPA vs 低GPA學生比較")
            try:
                gpa_list = []
                for idx, row in self.data.iterrows():
                    scores = [row['一般必修'], row['一般選修'], row['通識必修'], row['通識選修']]
                    valid_scores = [s for s in scores if pd.notna(s)]
                    if len(valid_scores) >= 2:
                        gpa_list.append((idx, np.mean(valid_scores)))
                
                if len(gpa_list) >= 20:
                    gpa_list.sort(key=lambda x: x[1])
                    n = len(gpa_list)
                    low_30_percent = int(n * 0.3)
                    high_30_percent = int(n * 0.7)
                    
                    low_gpa_indices = [x[0] for x in gpa_list[:low_30_percent]]
                    high_gpa_indices = [x[0] for x in gpa_list[high_30_percent:]]
                    
                    subjects = ['一般必修', '一般選修', '通識必修', '通識選修']
                    for subject in subjects:
                        high_scores = [self.data.loc[i, subject] for i in high_gpa_indices 
                                      if pd.notna(self.data.loc[i, subject])]
                        low_scores = [self.data.loc[i, subject] for i in low_gpa_indices 
                                     if pd.notna(self.data.loc[i, subject])]
                        
                        if len(high_scores) >= 2 and len(low_scores) >= 2:
                            statistic, p_value = stats.ttest_ind(high_scores, low_scores)
                            result_key = f"獨立樣本t-test_高GPA_vs_低GPA_{subject}"
                            all_results[result_key] = {
                                'type': 'independent_ttest',
                                'comparison': f"高GPA vs 低GPA ({subject})",
                                'statistic': statistic,
                                'p_value': p_value,
                                'mean1': float(np.mean(high_scores)),
                                'std1': float(np.std(high_scores)),
                                'n1': len(high_scores),
                                'mean2': float(np.mean(low_scores)),
                                'std2': float(np.std(low_scores)),
                                'n2': len(low_scores),
                                'mean_diff': float(np.mean(high_scores) - np.mean(low_scores)),
                                'significance': self._get_significance(p_value)
                            }
            except Exception as e:
                logger.error(f"分析高GPA vs 低GPA時發生錯誤: {str(e)}")
            
            # 6d. 必修高分學生的選修表現
            current_step += 1
            if progress_callback:
                progress_callback(current_step, "必修高分學生的選修課表現分析")
            try:
                required_avg = []
                for idx, row in self.data.iterrows():
                    req_scores = [row['一般必修'], row['通識必修']]
                    valid_req = [s for s in req_scores if pd.notna(s)]
                    if valid_req:
                        required_avg.append((idx, np.mean(valid_req)))
                
                if len(required_avg) >= 10:
                    required_avg.sort(key=lambda x: x[1], reverse=True)
                    top_30_percent = int(len(required_avg) * 0.3)
                    high_required_students = [x[0] for x in required_avg[:top_30_percent]]
                    
                    elective_scores = []
                    for idx in high_required_students:
                        row = self.data.loc[idx]
                        elec_scores = [row['一般選修'], row['通識選修']]
                        valid_elec = [s for s in elec_scores if pd.notna(s)]
                        if valid_elec:
                            elective_scores.append(np.mean(valid_elec))
                    
                    overall_elective = []
                    for idx, row in self.data.iterrows():
                        elec_scores = [row['一般選修'], row['通識選修']]
                        valid_elec = [s for s in elec_scores if pd.notna(s)]
                        if valid_elec:
                            overall_elective.append(np.mean(valid_elec))
                    
                    if len(elective_scores) >= 2 and len(overall_elective) >= 2:
                        statistic, p_value = stats.ttest_ind(elective_scores, overall_elective)
                        all_results["配對t-test_必修高分學生選修表現_vs_整體選修表現"] = {
                            'type': 'independent_ttest',
                            'comparison': "必修高分學生選修表現 vs 整體選修表現",
                            'statistic': statistic,
                            'p_value': p_value,
                            'mean1': float(np.mean(elective_scores)),
                            'std1': float(np.std(elective_scores)),
                            'n1': len(elective_scores),
                            'mean2': float(np.mean(overall_elective)),
                            'std2': float(np.std(overall_elective)),
                            'n2': len(overall_elective),
                            'mean_diff': float(np.mean(elective_scores) - np.mean(overall_elective)),
                            'significance': self._get_significance(p_value)
                        }
            except Exception as e:
                logger.error(f"分析必修高分學生選修表現時發生錯誤: {str(e)}")

            # ========== 第七類：學院間比較分析（獨立樣本t-test）==========
            # 目的：比較不同學院學生在各課程類型的學習表現差異
            colleges = ["理學院", "工學院", "商學院", "設計學院", "人文與教育學院", "法學院", "電機資訊學院"]
            course_types = ["一般必修", "一般選修", "通識必修", "通識選修"]
            
            logger.info("執行學院間比較分析")
            
            for i, college1 in enumerate(colleges):
                for j, college2 in enumerate(colleges):
                    if i < j:  # 避免重複比較
                        for course_type in course_types:
                            if self.operation_cancelled:
                                logger.info("操作被取消")
                                return all_results
                            
                            current_step += 1
                            if progress_callback:
                                progress_callback(current_step, f"學院比較: {college1} vs {college2} ({course_type})")
                            
                            try:
                                logger.debug(f"分析 {college1} vs {college2} ({course_type})")
                                college1_data = []
                                college2_data = []
                                
                                for idx, row in self.data.iterrows():
                                    college = row['學院']
                                    score = row[course_type]
                                    
                                    if pd.isna(score):
                                        continue
                                    
                                    if college == college1:
                                        college1_data.append(score)
                                    elif college == college2:
                                        college2_data.append(score)
                                
                                if len(college1_data) >= 2 and len(college2_data) >= 2:
                                    statistic, p_value = stats.ttest_ind(college1_data, college2_data)
                                    
                                    result_key = f"獨立樣本t-test_{college1}_vs_{college2}_{course_type}"
                                    all_results[result_key] = {
                                        'type': 'independent_ttest',
                                        'comparison': f"{college1} vs {college2} ({course_type})",
                                        'statistic': statistic,
                                        'p_value': p_value,
                                        'mean1': np.mean(college1_data),
                                        'std1': np.std(college1_data),
                                        'n1': len(college1_data),
                                        'mean2': np.mean(college2_data),
                                        'std2': np.std(college2_data),
                                        'n2': len(college2_data),
                                        'mean_diff': np.mean(college1_data) - np.mean(college2_data),
                                        'significance': self._get_significance(p_value)
                                    }
                                    logger.debug(f"完成 {college1} vs {college2} ({course_type}), p={p_value:.4f}")
                                else:
                                    logger.warning(f"{college1} vs {college2} ({course_type}): 資料不足 "
                                                 f"({len(college1_data)}, {len(college2_data)})")
                            except Exception as e:
                                logger.error(f"分析 {college1} vs {college2} ({course_type}) 時發生錯誤: {str(e)}")
                                continue
            
            return all_results
            
        except Exception as e:
            print(f"分析過程中發生錯誤: {str(e)}")
            return all_results
    
    def _get_significance(self, p_value):
        """判斷顯著性"""
        if p_value < 0.001:
            return "極顯著 (p < 0.001)"
        elif p_value < 0.01:
            return "高度顯著 (p < 0.01)"
        elif p_value < 0.05:
            return "顯著 (p < 0.05)"
        else:
            return "不顯著 (p >= 0.05)"
    
    def export_to_excel(self):
        """導出完整分析結果到Excel"""
        if self.data is None:
            messagebox.showwarning("警告", "請先載入資料檔案")
            return
        
        logger.info("開始導出完整分析報表")
        
        try:
            from tkinter import filedialog
            import datetime
            
            # 計算預估分析數量並建立進度視窗
            estimated_steps = 90  # 6個配對測試 + 84個學院比較
            self.create_progress_window("執行所有統計分析...", estimated_steps)
            
            # 執行所有分析
            all_results = self.run_all_analyses(progress_callback=self.update_progress)
            
            if self.operation_cancelled:
                self.close_progress_window()
                logger.info("用戶取消分析操作")
                return
            
            if not all_results:
                self.close_progress_window()
                messagebox.showwarning("警告", "無法執行分析，請檢查資料格式")
                logger.warning("分析結果為空")
                return
            
            # 選擇儲存位置
            self.close_progress_window()
            logger.info(f"完成分析，共 {len(all_results)} 項結果")
            
            filename = filedialog.asksaveasfilename(
                title="儲存完整分析報表",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"完整T-test分析報表_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            
            if not filename:
                logger.info("用戶取消檔案儲存")
                return
            
            # 建立Excel儲存進度視窗
            self.create_progress_window("儲存Excel檔案...", 7)
            logger.info(f"開始儲存Excel檔案: {filename}")
            
            # 建立Excel工作簿
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                
                # 1. 總覽工作表
                self.update_progress(1, "建立分析總覽...")
                overview_data = {
                    '項目': ['分析時間', '總分析數量', '資料筆數', '顯著結果數'],
                    '內容': [
                        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        len(all_results),
                        len(self.data),
                        sum(1 for result in all_results.values() if result['p_value'] < 0.05)
                    ]
                }
                overview_df = pd.DataFrame(overview_data)
                overview_df.to_excel(writer, sheet_name='分析總覽', index=False)
                logger.debug("完成分析總覽工作表")
                
                # 2. 配對t-test結果
                self.update_progress(2, "整理配對t-test結果...")
                paired_data = []
                for key, result in all_results.items():
                    if result['type'] == 'paired_ttest':
                        paired_data.append({
                            '比較項目': result['comparison'],
                            '組別1平均': f"{result['mean1']:.2f}",
                            '組別1標準差': f"{result['std1']:.2f}",
                            '組別2平均': f"{result['mean2']:.2f}",
                            '組別2標準差': f"{result['std2']:.2f}",
                            '平均差值': f"{result['mean_diff']:.2f}",
                            '樣本配對數': result['n_pairs'],
                            't統計量': f"{result['statistic']:.4f}",
                            'p值': f"{result['p_value']:.4f}",
                            '顯著性': result['significance']
                        })
                
                if paired_data:
                    paired_df = pd.DataFrame(paired_data)
                    paired_df.to_excel(writer, sheet_name='配對t-test結果', index=False)
                    logger.debug(f"完成配對t-test結果工作表，{len(paired_data)}項結果")
                
                # 3. 獨立樣本t-test結果
                independent_data = []
                for key, result in all_results.items():
                    if result['type'] == 'independent_ttest':
                        independent_data.append({
                            '比較項目': result['comparison'],
                            '組別1平均': f"{result['mean1']:.2f}",
                            '組別1標準差': f"{result['std1']:.2f}",
                            '組別1樣本數': result['n1'],
                            '組別2平均': f"{result['mean2']:.2f}",
                            '組別2標準差': f"{result['std2']:.2f}",
                            '組別2樣本數': result['n2'],
                            '平均差值': f"{result['mean_diff']:.2f}",
                            't統計量': f"{result['statistic']:.4f}",
                            'p值': f"{result['p_value']:.4f}",
                            '顯著性': result['significance']
                        })
                
                if independent_data:
                    independent_df = pd.DataFrame(independent_data)
                    independent_df.to_excel(writer, sheet_name='獨立樣本t-test結果', index=False)
                
                # 4. 顯著結果摘要
                significant_data = []
                for key, result in all_results.items():
                    if result['p_value'] < 0.05:
                        significant_data.append({
                            '分析類型': '配對t-test' if result['type'] == 'paired_ttest' else '獨立樣本t-test',
                            '比較項目': result['comparison'],
                            't統計量': f"{result['statistic']:.4f}",
                            'p值': f"{result['p_value']:.4f}",
                            '顯著性': result['significance'],
                            '效果方向': '組別1 > 組別2' if result['mean_diff'] > 0 else '組別1 < 組別2'
                        })
                
                if significant_data:
                    significant_df = pd.DataFrame(significant_data)
                    significant_df.to_excel(writer, sheet_name='顯著結果摘要', index=False)
                
                # 5. 原始資料範例
                sample_data = self.data.head(500)  # 限制為500筆以控制檔案大小
                sample_data.to_excel(writer, sheet_name='原始資料範例', index=False)
                
                # 6. 資料摘要統計
                summary_stats = self.data[['一般必修', '一般選修', '通識必修', '通識選修']].describe()
                summary_stats.to_excel(writer, sheet_name='資料摘要統計')

                # 7. 當前分析（若有）
                if self.current_analysis_result is not None:
                    try:
                        current_df = pd.DataFrame([
                            {
                                '分析標題': self.current_analysis_result.get('title', ''),
                                '顯著性': self.current_analysis_result.get('significance', ''),
                                't統計量': f"{self.current_analysis_result.get('statistic', np.nan):.4f}",
                                'p值': f"{self.current_analysis_result.get('p_value', np.nan):.4f}",
                                '時間戳': self.current_analysis_result.get('timestamp', '')
                            }
                        ])
                        current_df.to_excel(writer, sheet_name='當前分析', index=False)
                    except Exception:
                        pass
            
            self.update_progress(7, "完成!")
            self.close_progress_window()
            
            # 統計顯著結果
            significant_count = sum(1 for result in all_results.values() if result['p_value'] < 0.05)
            
            success_msg = f"完整分析報表已儲存至: {filename}\n\n" \
                         f"總共完成 {len(all_results)} 項分析\n" \
                         f"其中 {significant_count} 項達到顯著水準 (p < 0.05)"
            
            messagebox.showinfo("成功", success_msg)
            logger.info(f"Excel報表儲存完成: {filename}")
            logger.info(f"統計摘要: 總分析{len(all_results)}項, 顯著{significant_count}項")
            
        except Exception as e:
            self.close_progress_window()
            error_msg = f"導出Excel時發生錯誤: {str(e)}"
            logger.error(error_msg)
            logger.error(f"錯誤詳情: {traceback.format_exc()}")
            messagebox.showerror("錯誤", error_msg)
    
    def clear_results(self):
        """清空分析結果"""
        self.result_text.delete(1.0, tk.END)
        self.current_analysis_result = None
        messagebox.showinfo("完成", "分析結果已清空")


def main():
    logger.info("=" * 50)
    logger.info("T-test分析工具啟動")
    logger.info(f"Python版本: {sys.version}")
    logger.info(f"Pandas版本: {pd.__version__}")
    logger.info(f"NumPy版本: {np.__version__}")
    logger.info(f"SciPy版本: {stats.__version__ if hasattr(stats, '__version__') else 'Unknown'}")
    logger.info(f"工作目錄: {os.getcwd()}")
    logger.info("=" * 50)
    
    try:
        root = tk.Tk()
        app = TTestAnalyzer(root)
        logger.info("開始GUI主迴圈")
        root.mainloop()
        logger.info("程式正常結束")
    except Exception as e:
        logger.error(f"程式執行時發生嚴重錯誤: {str(e)}")
        logger.error(f"錯誤詳情: {traceback.format_exc()}")
        raise

if __name__ == "__main__":
    main()