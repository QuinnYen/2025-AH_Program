import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import pearsonr
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 設定中文字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False

class CorrelationAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("中原大學學生成績相關性分析")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 檔案路徑
        self.file_path = tk.StringVar()
        
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 標題
        title_label = ttk.Label(main_frame, text="中原大學學生成績相關性分析 (Excel版)", 
                               font=("Microsoft JhengHei", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 檔案選擇區域
        file_frame = ttk.LabelFrame(main_frame, text="步驟1: 選擇資料檔案", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(file_frame, text="選擇Excel檔案:").grid(row=0, column=0, sticky=tk.W)
        
        # 設定預設路徑為專案根目錄
        default_path = os.getcwd()
        self.file_path.set(default_path)
        
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(file_frame, text="瀏覽...", command=self.browse_file).grid(row=1, column=1)
        
        file_frame.columnconfigure(0, weight=1)
        
        # 分析選項區域
        options_frame = ttk.LabelFrame(main_frame, text="步驟2: 選擇分析選項", padding="10")
        options_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.create_heatmap = tk.BooleanVar(value=True)
        self.create_scatter = tk.BooleanVar(value=True)
        self.analyze_by_college = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(options_frame, text="產生相關性熱力圖", variable=self.create_heatmap).grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="產生散佈圖矩陣", variable=self.create_scatter).grid(row=1, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="按學院分析", variable=self.analyze_by_college).grid(row=2, column=0, sticky=tk.W)
        
        # 執行按鈕
        self.analyze_button = ttk.Button(main_frame, text="開始分析", command=self.start_analysis, style="Accent.TButton")
        self.analyze_button.grid(row=3, column=0, columnspan=2, pady=20)
        
        # 進度條
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 狀態標籤
        self.status_label = ttk.Label(main_frame, text="預設路徑已設定為專案根目錄，請選擇Excel檔案開始分析")
        self.status_label.grid(row=5, column=0, columnspan=2)
        
        # 結果顯示區域
        results_frame = ttk.LabelFrame(main_frame, text="分析結果", padding="10")
        results_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 結果文本框
        self.results_text = tk.Text(results_frame, height=15, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # 設定主視窗的權重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def browse_file(self):
        # 設定初始目錄為專案根目錄
        initial_dir = os.getcwd()
        
        filename = filedialog.askopenfilename(
            title="選擇Excel檔案",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()
        
    def update_results(self, message):
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.see(tk.END)
        self.root.update()
        
    def clear_results(self):
        self.results_text.delete(1.0, tk.END)
        
    def start_analysis(self):
        file_path = self.file_path.get()
        
        # 檢查是否選擇了具體的檔案
        if not file_path or os.path.isdir(file_path):
            messagebox.showerror("錯誤", "請選擇具體的Excel檔案!")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("錯誤", "選擇的檔案不存在!")
            return
            
        # 在新執行緒中執行分析，避免GUI卡住
        analysis_thread = threading.Thread(target=self.run_analysis)
        analysis_thread.daemon = True
        analysis_thread.start()
        
    def run_analysis(self):
        try:
            self.analyze_button.config(state='disabled')
            self.progress.start()
            self.clear_results()
            
            # 執行分析
            self.perform_analysis()
            
        except Exception as e:
            messagebox.showerror("分析錯誤", f"分析過程中發生錯誤:\n{str(e)}")
        finally:
            self.progress.stop()
            self.analyze_button.config(state='normal')
            
    def perform_analysis(self):
        file_path = self.file_path.get()
        
        # 檢查是否為檔案路徑還是目錄路徑
        if os.path.isdir(file_path):
            messagebox.showerror("錯誤", "請選擇Excel檔案，而非目錄!")
            return
            
        if not file_path.endswith(('.xlsx', '.xls')):
            messagebox.showerror("錯誤", "請選擇Excel檔案!")
            return
        
        # 1. 載入資料
        self.update_status("正在載入資料...")
        self.update_results("=== 開始分析 ===")
        self.update_results(f"檔案路徑: {file_path}")
        
        # 讀取Excel檔案
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            messagebox.showerror("讀取錯誤", f"無法讀取Excel檔案:\n{str(e)}")
            return
            
        self.update_results(f"原始資料筆數: {len(df):,}")
        
        # 檢查必要欄位是否存在
        score_columns = ['一般必修', '一般選修', '通識必修', '通識選修']
        missing_columns = [col for col in score_columns if col not in df.columns]
        if missing_columns:
            messagebox.showerror("資料格式錯誤", f"檔案中缺少以下欄位:\n{', '.join(missing_columns)}")
            return
        
        # 清理資料
        df_clean = df.dropna(subset=score_columns)
        self.update_results(f"有效資料筆數: {len(df_clean):,}")
        self.update_results(f"移除無效資料: {len(df) - len(df_clean):,}")
        
        # 2. 基本統計
        self.update_status("計算基本統計...")
        self.update_results("\n=== 基本統計資訊 ===")
        for col in score_columns:
            mean_score = df_clean[col].mean()
            std_score = df_clean[col].std()
            self.update_results(f"{col}: 平均 {mean_score:.2f} ± {std_score:.2f}")
        
        # 3. 相關性分析
        self.update_status("計算相關性...")
        self.update_results("\n=== 相關性分析結果 ===")
        
        correlation_matrix = df_clean[score_columns].corr()
        
        # 詳細兩兩分析
        pairs = [
            ('一般必修', '一般選修'),
            ('一般必修', '通識必修'),
            ('一般必修', '通識選修'),
            ('一般選修', '通識必修'),
            ('一般選修', '通識選修'),
            ('通識必修', '通識選修')
        ]
        
        results_data = []
        
        for var1, var2 in pairs:
            paired_data = df_clean[[var1, var2]].dropna()
            corr_coef = paired_data[var1].corr(paired_data[var2])
            
            # 判斷相關強度
            if abs(corr_coef) >= 0.7:
                strength = "強相關"
            elif abs(corr_coef) >= 0.3:
                strength = "中等相關"
            else:
                strength = "弱相關"
            
            # 計算p值
            _, p_value = pearsonr(paired_data[var1], paired_data[var2])
            
            if p_value < 0.001:
                significance = "極顯著 ***"
            elif p_value < 0.01:
                significance = "很顯著 **"
            elif p_value < 0.05:
                significance = "顯著 *"
            else:
                significance = "不顯著"
            
            self.update_results(f"{var1} ↔ {var2}: r = {corr_coef:.3f} ({strength}, {significance})")
            
            results_data.append({
                '變數1': var1,
                '變數2': var2,
                '樣本數': len(paired_data),
                '相關係數': corr_coef,
                'p值': p_value,
                '相關強度': strength,
                '顯著性': significance
            })
        
        # 4. 按學院分析 (如果選擇)
        if self.analyze_by_college.get() and '學院' in df_clean.columns:
            self.update_status("按學院分析...")
            self.update_results("\n=== 按學院分析 ===")
            
            colleges = df_clean['學院'].unique()
            for college in colleges:
                college_data = df_clean[df_clean['學院'] == college]
                if len(college_data) > 30:
                    # 計算學院內最高相關性
                    college_corr = college_data[score_columns].corr()
                    max_corr = college_corr.abs().where(np.triu(np.ones(college_corr.shape), k=1).astype(bool)).max().max()
                    self.update_results(f"{college} (n={len(college_data)}): 最高相關性 = {max_corr:.3f}")
        
        # 5. 產生視覺化
        output_dir = os.path.dirname(file_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if self.create_heatmap.get():
            self.update_status("產生熱力圖...")
            self.create_heatmap_chart(correlation_matrix, os.path.join(output_dir, f"correlation_heatmap_{timestamp}.png"))
            
        if self.create_scatter.get():
            self.update_status("產生散佈圖...")
            self.create_scatter_chart(df_clean, pairs, os.path.join(output_dir, f"scatter_plots_{timestamp}.png"))
        
        # 6. 匯出Excel結果
        self.update_status("匯出結果...")
        excel_path = os.path.join(output_dir, f"相關性分析結果_{timestamp}.xlsx")
        self.export_to_excel(pd.DataFrame(results_data), correlation_matrix, excel_path)
        
        # 完成
        self.update_results("\n=== 分析完成 ===")
        self.update_results(f"結果檔案已儲存至: {output_dir}")
        self.update_results(f"Excel檔案: 相關性分析結果_{timestamp}.xlsx")
        
        if self.create_heatmap.get():
            self.update_results(f"熱力圖: correlation_heatmap_{timestamp}.png")
        if self.create_scatter.get():
            self.update_results(f"散佈圖: scatter_plots_{timestamp}.png")
            
        self.update_status("分析完成!")
        messagebox.showinfo("完成", f"分析完成!\n結果已儲存至:\n{output_dir}")
        
    def create_heatmap_chart(self, correlation_matrix, save_path):
        plt.figure(figsize=(10, 8))
        mask = np.triu(np.ones_like(correlation_matrix, dtype=bool))
        sns.heatmap(correlation_matrix, 
                    mask=mask,
                    annot=True, 
                    cmap='RdYlBu_r', 
                    center=0,
                    square=True,
                    fmt='.3f',
                    cbar_kws={"shrink": .8})
        
        plt.title('學生成績相關性分析熱力圖', fontsize=16, fontweight='bold', pad=20)
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
    def create_scatter_chart(self, df, pairs, save_path):
        fig, axes = plt.subplots(2, 3, figsize=(18, 12))
        fig.suptitle('學生成績相關性散佈圖', fontsize=16, fontweight='bold')
        axes = axes.flatten()
        
        for idx, (var1, var2) in enumerate(pairs):
            plot_data = df[[var1, var2]].dropna()
            if len(plot_data) > 0:
                corr_coef = plot_data[var1].corr(plot_data[var2])
                axes[idx].scatter(plot_data[var1], plot_data[var2], alpha=0.5, s=1)
                
                # 趨勢線
                z = np.polyfit(plot_data[var1], plot_data[var2], 1)
                p = np.poly1d(z)
                axes[idx].plot(plot_data[var1], p(plot_data[var1]), "r--", alpha=0.8)
                
                axes[idx].set_xlabel(var1)
                axes[idx].set_ylabel(var2)
                axes[idx].set_title(f'{var1} vs {var2}\nr = {corr_coef:.3f}')
                axes[idx].grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
    def export_to_excel(self, results_df, correlation_matrix, file_path):
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='詳細相關性分析', index=False)
            correlation_matrix.to_excel(writer, sheet_name='相關性矩陣')
            
            # 解釋說明
            interpretation = pd.DataFrame({
                '相關係數範圍': ['0.7 ≤ |r| ≤ 1.0', '0.3 ≤ |r| < 0.7', '0.0 ≤ |r| < 0.3'],
                '相關強度': ['強相關', '中等相關', '弱相關'],
                '教育意義': [
                    '學生在這兩類課程表現高度一致，可互相預測',
                    '存在中等程度關聯，但仍有個別差異', 
                    '兩類課程評估不同能力，關聯性很低'
                ]
            })
            interpretation.to_excel(writer, sheet_name='結果解釋', index=False)

def main():
    root = tk.Tk()
    app = CorrelationAnalysisGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()