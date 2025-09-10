import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import pearsonr
from scipy import stats
from sklearn.preprocessing import StandardScaler
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
        self.gpa_stratified_analysis = tk.BooleanVar(value=False)
        self.partial_correlation_analysis = tk.BooleanVar(value=False)
        self.longitudinal_analysis = tk.BooleanVar(value=False)
        
        # 基礎分析選項
        ttk.Label(options_frame, text="基礎分析：", font=("Microsoft JhengHei", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0,5))
        ttk.Checkbutton(options_frame, text="產生相關性熱力圖", variable=self.create_heatmap).grid(row=1, column=0, sticky=tk.W, padx=(20,0))
        ttk.Checkbutton(options_frame, text="產生散佈圖矩陣", variable=self.create_scatter).grid(row=2, column=0, sticky=tk.W, padx=(20,0))
        ttk.Checkbutton(options_frame, text="按學院分析", variable=self.analyze_by_college).grid(row=3, column=0, sticky=tk.W, padx=(20,0))
        
        # 進階分析選項
        ttk.Label(options_frame, text="進階分析：", font=("Microsoft JhengHei", 10, "bold")).grid(row=4, column=0, sticky=tk.W, pady=(10,5))
        ttk.Checkbutton(options_frame, text="GPA分層學習連結分析", variable=self.gpa_stratified_analysis).grid(row=5, column=0, sticky=tk.W, padx=(20,0))
        ttk.Checkbutton(options_frame, text="必修課預測能力分析（偏相關）", variable=self.partial_correlation_analysis).grid(row=6, column=0, sticky=tk.W, padx=(20,0))
        ttk.Checkbutton(options_frame, text="學習軌跡縱向分析（需多學年檔案）", variable=self.longitudinal_analysis).grid(row=7, column=0, sticky=tk.W, padx=(20,0))
        
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
        
        # 4. GPA分層學習連結分析
        gpa_stratified_results = []
        if self.gpa_stratified_analysis.get():
            self.update_status("執行GPA分層學習連結分析...")
            gpa_stratified_results = self.perform_gpa_stratified_analysis(df_clean, score_columns)
        
        # 5. 必修課預測能力分析（偏相關）
        partial_corr_results = []
        if self.partial_correlation_analysis.get():
            self.update_status("執行必修課預測能力分析...")
            partial_corr_results = self.perform_partial_correlation_analysis(df_clean, score_columns)
        
        # 5.5. 學習軌跡縱向分析
        longitudinal_results = []
        if self.longitudinal_analysis.get():
            self.update_status("執行學習軌跡縱向分析...")
            longitudinal_results = self.perform_longitudinal_analysis(file_path, df_clean, score_columns)
        
        # 6. 詳細學院相關性分析 (如果選擇)
        college_results = []  # 儲存學院分析結果
        college_detailed_correlations = []  # 儲存詳細的學院相關性資料
        if self.analyze_by_college.get() and '學院' in df_clean.columns:
            self.update_status("執行詳細學院相關性分析...")
            college_results, college_detailed_correlations = self.perform_detailed_college_analysis(df_clean, score_columns)
            
            # 增強學院課程結構關聯分析
            if len(college_results) >= 2:
                self.enhanced_college_structure_analysis(df_clean, score_columns, college_results)
        
        # 5. 產生視覺化
        output_dir = os.path.dirname(file_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if self.create_heatmap.get():
            self.update_status("產生熱力圖...")
            self.create_heatmap_chart(correlation_matrix, os.path.join(output_dir, f"correlation_heatmap_{timestamp}.png"))
            
        if self.create_scatter.get():
            self.update_status("產生散佈圖...")
            self.create_scatter_chart(df_clean, pairs, os.path.join(output_dir, f"scatter_plots_{timestamp}.png"))
        
        # 7. 匯出Excel結果
        self.update_status("匯出結果...")
        excel_path = os.path.join(output_dir, f"相關性分析結果_{timestamp}.xlsx")
        self.export_to_excel(pd.DataFrame(results_data), correlation_matrix, excel_path, 
                           college_results, gpa_stratified_results, partial_corr_results, 
                           longitudinal_results, college_detailed_correlations)
        
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
        
    def export_to_excel(self, results_df, correlation_matrix, file_path, college_results=None, 
                       gpa_stratified_results=None, partial_corr_results=None, 
                       longitudinal_results=None, college_detailed_correlations=None):
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            results_df.to_excel(writer, sheet_name='詳細相關性分析', index=False)
            correlation_matrix.to_excel(writer, sheet_name='相關性矩陣')
            
            # 新增：匯出學院分析結果
            if college_results:
                college_df = pd.DataFrame(college_results)
                college_df.to_excel(writer, sheet_name='學院分析摘要', index=False)
            
            # 新增：匯出詳細學院相關性分析
            if college_detailed_correlations:
                college_detailed_df = pd.DataFrame(college_detailed_correlations)
                college_detailed_df.to_excel(writer, sheet_name='各學院詳細相關性', index=False)
            
            # 新增：匯出GPA分層分析結果
            if gpa_stratified_results:
                # 分層相關性結果
                gpa_corr_df = pd.DataFrame(gpa_stratified_results['correlations'])
                gpa_corr_df.to_excel(writer, sheet_name='GPA分層相關性', index=False)
                
                # 分層統計摘要
                gpa_summary_df = pd.DataFrame(gpa_stratified_results['summary'])
                gpa_summary_df.to_excel(writer, sheet_name='GPA分層統計摘要', index=False)
            
            # 新增：匯出偏相關分析結果
            if partial_corr_results:
                partial_df = pd.DataFrame(partial_corr_results)
                partial_df.to_excel(writer, sheet_name='偏相關分析', index=False)
            
            # 新增：匯出縱向分析結果
            if longitudinal_results:
                longitudinal_df = pd.DataFrame(longitudinal_results)
                longitudinal_df.to_excel(writer, sheet_name='學習軌跡縱向分析', index=False)
            
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
    
    def perform_gpa_stratified_analysis(self, df, score_columns):
        """
        GPA分層學習連結分析
        將學生分為高分組(前30%)、中分組(中40%)、低分組(後30%)三群
        分別計算這三組學生的課程相關性矩陣
        """
        self.update_results("\n=== GPA分層學習連結分析 ===")
        
        # 計算總體GPA
        df['總體GPA'] = df[score_columns].mean(axis=1)
        
        # 分層
        n = len(df)
        high_threshold = df['總體GPA'].quantile(0.7)  # 前30%
        low_threshold = df['總體GPA'].quantile(0.3)   # 後30%
        
        high_group = df[df['總體GPA'] >= high_threshold]
        medium_group = df[(df['總體GPA'] < high_threshold) & (df['總體GPA'] > low_threshold)]
        low_group = df[df['總體GPA'] <= low_threshold]
        
        groups = {
            '高分組': high_group,
            '中分組': medium_group, 
            '低分組': low_group
        }
        
        correlations = []
        summary = []
        
        for group_name, group_data in groups.items():
            self.update_results(f"\n{group_name} (n={len(group_data)}):")
            
            # 計算組內相關性矩陣
            group_corr = group_data[score_columns].corr()
            
            # 計算平均相關性（排除對角線）
            mask = np.triu(np.ones_like(group_corr, dtype=bool), k=1)
            correlations_list = group_corr.values[mask]
            avg_correlation = np.mean(correlations_list)
            
            # 記錄詳細相關性
            pairs = [('一般必修', '一般選修'), ('一般必修', '通識必修'), 
                    ('一般必修', '通識選修'), ('一般選修', '通識必修'),
                    ('一般選修', '通識選修'), ('通識必修', '通識選修')]
            
            for var1, var2 in pairs:
                corr_value = group_corr.loc[var1, var2]
                correlations.append({
                    '群組': group_name,
                    '變數1': var1,
                    '變數2': var2,
                    '相關係數': corr_value,
                    '樣本數': len(group_data)
                })
                self.update_results(f"  {var1} ↔ {var2}: r = {corr_value:.3f}")
            
            # 學習連貫性評估
            if avg_correlation >= 0.6:
                coherence = "高度連貫（一好俱好）"
            elif avg_correlation >= 0.3:
                coherence = "中等連貫"
            else:
                coherence = "低連貫性（各科目相對獨立）"
            
            summary.append({
                '群組': group_name,
                '樣本數': len(group_data),
                'GPA範圍': f"{group_data['總體GPA'].min():.2f} - {group_data['總體GPA'].max():.2f}",
                '平均相關性': avg_correlation,
                '學習連貫性': coherence,
                '最高相關性': np.max(correlations_list),
                '最低相關性': np.min(correlations_list)
            })
            
            self.update_results(f"  平均相關性: {avg_correlation:.3f} ({coherence})")
        
        # 結論
        self.update_results("\n【GPA分層分析結論】")
        high_coherence = summary[0]['平均相關性']
        low_coherence = summary[2]['平均相關性']
        
        if high_coherence > low_coherence + 0.1:
            self.update_results("高分組學生確實展現「一好俱好」的學習連貫性")
            self.update_results("低分組學生的課程表現較為分散，各科目相對獨立")
        else:
            self.update_results("不同GPA層級學生的課程相關性模式相近")
        
        return {
            'correlations': correlations,
            'summary': summary
        }
    
    def perform_partial_correlation_analysis(self, df, score_columns):
        """
        必修課預測能力分析（偏相關分析）
        分析「一般必修」與「一般選修」的相關性時，排除「通識課程」成績的影響
        """
        self.update_results("\n=== 必修課預測能力分析（偏相關） ===")
        
        # 計算通識平均作為控制變數
        df['通識平均'] = df[['通識必修', '通識選修']].mean(axis=1)
        
        # 移除缺失值
        analysis_data = df[['一般必修', '一般選修', '通識平均']].dropna()
        
        if len(analysis_data) < 30:
            self.update_results("資料不足，無法進行偏相關分析")
            return []
        
        # 計算簡單相關性（未控制前）
        simple_corr = analysis_data['一般必修'].corr(analysis_data['一般選修'])
        
        # 計算偏相關係數
        # 公式: r_xy.z = (r_xy - r_xz * r_yz) / sqrt((1-r_xz²)(1-r_yz²))
        r_xy = simple_corr
        r_xz = analysis_data['一般必修'].corr(analysis_data['通識平均'])
        r_yz = analysis_data['一般選修'].corr(analysis_data['通識平均'])
        
        numerator = r_xy - (r_xz * r_yz)
        denominator = np.sqrt((1 - r_xz**2) * (1 - r_yz**2))
        partial_corr = numerator / denominator if denominator != 0 else 0
        
        # 計算統計顯著性（偏相關的t檢驗）
        n = len(analysis_data)
        df_partial = n - 3  # 自由度 = n - k - 1，其中k=1個控制變數
        t_stat = partial_corr * np.sqrt(df_partial / (1 - partial_corr**2))
        p_value = 2 * (1 - stats.t.cdf(abs(t_stat), df_partial))
        
        results = [{
            '分析項目': '一般必修 vs 一般選修',
            '簡單相關係數': simple_corr,
            '偏相關係數': partial_corr,
            '控制變數': '通識課程平均',
            '樣本數': n,
            't統計量': t_stat,
            'p值': p_value,
            '顯著性': '顯著' if p_value < 0.05 else '不顯著',
            '效果解釋': '純粹的專業領域內關聯' if abs(partial_corr) > 0.3 else '控制學術能力後關聯微弱'
        }]
        
        self.update_results(f"簡單相關係數: r = {simple_corr:.3f}")
        self.update_results(f"偏相關係數: r_partial = {partial_corr:.3f} (控制通識成績)")
        self.update_results(f"統計顯著性: {'顯著' if p_value < 0.05 else '不顯著'} (p = {p_value:.4f})")
        
        # 結論
        diff = abs(simple_corr) - abs(partial_corr)
        if diff > 0.1:
            self.update_results("【結論】控制學術能力後，專業領域內關聯性明顯下降")
            self.update_results("必修課成績的預測能力主要來自學生整體學術能力")
        elif diff < -0.1:
            self.update_results("【結論】控制學術能力後，專業領域內關聯性反而增強")
            self.update_results("存在專業領域特有的學習關聯")
        else:
            self.update_results("【結論】控制學術能力前後，專業領域關聯性變化不大")
            self.update_results("必修與選修有獨立的預測關係")
        
        return results
    
    def enhanced_college_structure_analysis(self, df, score_columns, college_results):
        """
        增強的學院課程結構關聯分析
        比較不同學院的課程相關性模式差異
        """
        self.update_results("\n=== 學院課程結構深度分析 ===")
        
        # 重點比較設計學院和商學院
        target_colleges = ['設計學院', '商學院']
        available_colleges = [college['學院'] for college in college_results]
        
        for target_college in target_colleges:
            if target_college in available_colleges:
                college_data = df[df['學院'] == target_college]
                if len(college_data) > 20:
                    self.update_results(f"\n【{target_college}課程結構分析】")
                    
                    # 計算詳細相關性
                    college_corr = college_data[score_columns].corr()
                    
                    # 專業課程內部相關性
                    major_corr = college_corr.loc['一般必修', '一般選修']
                    general_corr = college_corr.loc['通識必修', '通識選修']
                    
                    # 跨領域相關性
                    cross_corr = np.mean([
                        college_corr.loc['一般必修', '通識必修'],
                        college_corr.loc['一般必修', '通識選修'],
                        college_corr.loc['一般選修', '通識必修'],
                        college_corr.loc['一般選修', '通識選修']
                    ])
                    
                    self.update_results(f"  專業課程內部關聯: {major_corr:.3f}")
                    self.update_results(f"  通識課程內部關聯: {general_corr:.3f}")
                    self.update_results(f"  跨領域平均關聯: {cross_corr:.3f}")
                    
                    # 課程結構特徵判斷
                    if major_corr < 0.3:
                        structure = "多元分化型（理論與實作差異大）"
                    elif major_corr > 0.6:
                        structure = "統合連貫型（課程設計一致性高）"
                    else:
                        structure = "平衡整合型"
                    
                    self.update_results(f"  課程結構特徵: {structure}")
        
        # 學院間比較
        if len(college_results) >= 2:
            self.update_results("\n【學院間課程關聯差異】")
            correlations = [college['最高相關性'] for college in college_results]
            max_corr_college = max(college_results, key=lambda x: x['最高相關性'])
            min_corr_college = min(college_results, key=lambda x: x['最高相關性'])
            
            self.update_results(f"最高課程關聯: {max_corr_college['學院']} ({max_corr_college['最高相關性']:.3f})")
            self.update_results(f"最低課程關聯: {min_corr_college['學院']} ({min_corr_college['最高相關性']:.3f})")
            
            if max_corr_college['最高相關性'] - min_corr_college['最高相關性'] > 0.2:
                self.update_results("不同學院確實存在顯著的課程結構差異")
            else:
                self.update_results("各學院課程結構相對一致")
    
    def perform_longitudinal_analysis(self, current_file_path, current_df, score_columns):
        """
        學習軌跡縱向分析
        嘗試找到並分析多個學年的資料，計算學習表現的穩定性
        """
        self.update_results("\n=== 學習軌跡縱向分析 ===")
        
        # 嘗試找到其他學年檔案
        base_dir = os.path.dirname(current_file_path)
        current_filename = os.path.basename(current_file_path)
        
        # 提取當前學年（假設檔案名包含學年資訊）
        year_files = []
        for file in os.listdir(base_dir):
            if file.endswith(('.xlsx', '.xls')) and '學年度' in file:
                year_files.append(os.path.join(base_dir, file))
        
        if len(year_files) < 2:
            self.update_results("未找到足夠的多學年檔案，無法進行縱向分析")
            self.update_results("建議將不同學年的資料檔案放在同一目錄下")
            return []
        
        self.update_results(f"發現 {len(year_files)} 個學年資料檔案")
        
        # 載入多個學年資料
        year_data = {}
        for file_path in year_files[:3]:  # 最多分析3個學年避免過度複雜
            try:
                filename = os.path.basename(file_path)
                # 提取學年資訊
                if '110' in filename:
                    year = '110學年'
                elif '111' in filename:
                    year = '111學年'
                elif '112' in filename:
                    year = '112學年'
                elif '113' in filename:
                    year = '113學年'
                else:
                    year = f"學年_{len(year_data)+1}"
                
                df = pd.read_excel(file_path)
                if all(col in df.columns for col in score_columns) and '學號' in df.columns:
                    # 只保留有完整資料的學生
                    df_clean = df.dropna(subset=score_columns + ['學號'])
                    year_data[year] = df_clean
                    self.update_results(f"  {year}: {len(df_clean)} 筆有效資料")
            except Exception as e:
                self.update_results(f"  無法載入 {filename}: {str(e)}")
        
        if len(year_data) < 2:
            self.update_results("可用的學年資料不足，無法進行縱向分析")
            return []
        
        # 找出跨學年共同學生
        year_list = list(year_data.keys())
        results = []
        
        for i in range(len(year_list)-1):
            year1 = year_list[i]
            year2 = year_list[i+1]
            
            # 找出兩學年都有資料的學生
            students1 = set(year_data[year1]['學號'])
            students2 = set(year_data[year2]['學號'])
            common_students = students1.intersection(students2)
            
            if len(common_students) < 20:
                self.update_results(f"{year1} vs {year2}: 共同學生不足({len(common_students)}人)，跳過")
                continue
            
            self.update_results(f"\n{year1} vs {year2} 縱向分析 (共同學生: {len(common_students)}人)")
            
            # 建立配對資料
            paired_data = []
            for student_id in common_students:
                try:
                    data1 = year_data[year1][year_data[year1]['學號'] == student_id].iloc[0]
                    data2 = year_data[year2][year_data[year2]['學號'] == student_id].iloc[0]
                    
                    paired_data.append({
                        '學號': student_id,
                        f'{year1}_必修平均': (data1['一般必修'] + data1['通識必修']) / 2,
                        f'{year1}_選修平均': (data1['一般選修'] + data1['通識選修']) / 2,
                        f'{year1}_總平均': data1[score_columns].mean(),
                        f'{year2}_必修平均': (data2['一般必修'] + data2['通識必修']) / 2,
                        f'{year2}_選修平均': (data2['一般選修'] + data2['通識選修']) / 2,
                        f'{year2}_總平均': data2[score_columns].mean(),
                    })
                except:
                    continue
            
            if len(paired_data) < 20:
                continue
                
            paired_df = pd.DataFrame(paired_data)
            
            # 計算縱向相關性
            correlations = {
                '必修平均穩定性': paired_df[f'{year1}_必修平均'].corr(paired_df[f'{year2}_必修平均']),
                '選修平均穩定性': paired_df[f'{year1}_選修平均'].corr(paired_df[f'{year2}_選修平均']),
                '總平均穩定性': paired_df[f'{year1}_總平均'].corr(paired_df[f'{year2}_總平均'])
            }
            
            # 記錄結果
            for measure, correlation in correlations.items():
                stability_level = "高穩定性" if correlation >= 0.7 else "中等穩定性" if correlation >= 0.5 else "低穩定性"
                results.append({
                    '比較學年': f"{year1} vs {year2}",
                    '測量指標': measure,
                    '相關係數': correlation,
                    '穩定性評估': stability_level,
                    '樣本數': len(paired_data),
                    '教育意義': self._interpret_stability(correlation, measure)
                })
                
                self.update_results(f"  {measure}: r = {correlation:.3f} ({stability_level})")
        
        # 整體結論
        if results:
            avg_stability = np.mean([r['相關係數'] for r in results])
            self.update_results(f"\n【縱向分析結論】")
            self.update_results(f"整體學習穩定性: {avg_stability:.3f}")
            
            if avg_stability >= 0.7:
                self.update_results("學生學業表現具有高度穩定性與可預測性")
            elif avg_stability >= 0.5:
                self.update_results("學業表現中等穩定，存在一定波動")
            else:
                self.update_results("學業表現波動較大，可能受課程難度或學習狀態影響")
        
        return results
    
    def _interpret_stability(self, correlation, measure):
        """解釋穩定性相關係數的教育意義"""
        if correlation >= 0.7:
            if '必修' in measure:
                return "必修課程表現高度穩定，反映基礎學力一致性"
            elif '選修' in measure:
                return "選修課程表現高度穩定，反映學習興趣與能力的持續性"
            else:
                return "整體學業表現高度可預測，學習模式穩定"
        elif correlation >= 0.5:
            return "表現具中等穩定性，受個別因素影響但整體趨勢一致"
        else:
            return "表現波動較大，可能受課程性質、教學方法或個人狀態影響"
    
    def perform_detailed_college_analysis(self, df, score_columns):
        """
        詳細學院相關性分析
        像 t-test 一樣，對每個學院進行完整的相關性分析
        """
        self.update_results("\n=== 詳細學院相關性分析 ===")
        
        colleges = df['學院'].unique()
        college_results = []
        detailed_correlations = []
        
        # 定義所有課程對組合
        pairs = [
            ('一般必修', '一般選修'),
            ('一般必修', '通識必修'),
            ('一般必修', '通識選修'),
            ('一般選修', '通識必修'),
            ('一般選修', '通識選修'),
            ('通識必修', '通識選修')
        ]
        
        for college in colleges:
            college_data = df[df['學院'] == college]
            
            # 只分析樣本數足夠的學院
            if len(college_data) < 20:  # 降低門檻讓更多學院參與分析
                self.update_results(f"{college}: 樣本數不足 (n={len(college_data)})，跳過分析")
                continue
            
            self.update_results(f"\n【{college}】(n={len(college_data)})")
            
            # 計算該學院的相關性矩陣
            college_corr = college_data[score_columns].corr()
            
            # 詳細分析每個課程對
            college_correlations = []
            for var1, var2 in pairs:
                # 取得有效資料
                paired_data = college_data[[var1, var2]].dropna()
                
                if len(paired_data) < 10:  # 配對資料太少則跳過
                    continue
                
                # 計算相關係數和統計顯著性
                corr_coef = paired_data[var1].corr(paired_data[var2])
                try:
                    _, p_value = pearsonr(paired_data[var1], paired_data[var2])
                except:
                    p_value = 1.0
                
                # 判斷相關強度
                if abs(corr_coef) >= 0.7:
                    strength = "強相關"
                elif abs(corr_coef) >= 0.3:
                    strength = "中等相關"
                else:
                    strength = "弱相關"
                
                # 判斷統計顯著性
                if p_value < 0.001:
                    significance = "極顯著***"
                elif p_value < 0.01:
                    significance = "很顯著**"
                elif p_value < 0.05:
                    significance = "顯著*"
                else:
                    significance = "不顯著"
                
                college_correlations.append(corr_coef)
                
                # 記錄詳細結果
                detailed_correlations.append({
                    '學院': college,
                    '課程對': f"{var1} ↔ {var2}",
                    '變數1': var1,
                    '變數2': var2,
                    '樣本數': len(paired_data),
                    '相關係數': corr_coef,
                    'p值': p_value,
                    '相關強度': strength,
                    '顯著性': significance,
                    '平均分1': paired_data[var1].mean(),
                    '標準差1': paired_data[var1].std(),
                    '平均分2': paired_data[var2].mean(),
                    '標準差2': paired_data[var2].std()
                })
                
                self.update_results(f"  {var1} ↔ {var2}: r = {corr_coef:.3f} ({strength}, {significance})")
            
            # 計算該學院的整體相關性統計
            if college_correlations:
                max_corr = max([abs(c) for c in college_correlations])
                min_corr = min([abs(c) for c in college_correlations])
                avg_corr = np.mean([abs(c) for c in college_correlations])
                
                # 學院課程結構特徵
                if avg_corr >= 0.6:
                    structure_type = "高度整合型"
                    interpretation = "各類課程高度關聯，知識結構統一"
                elif avg_corr >= 0.4:
                    structure_type = "中等整合型"
                    interpretation = "課程間存在中等關聯，部分知識互通"
                elif avg_corr >= 0.2:
                    structure_type = "低度整合型"
                    interpretation = "各課程相對獨立，專業分工明確"
                else:
                    structure_type = "分化專精型"
                    interpretation = "各課程高度分化，評估不同能力向度"
                
                college_results.append({
                    '學院': college,
                    '樣本數': len(college_data),
                    '最高相關性': max_corr,
                    '最低相關性': min_corr,
                    '平均相關性': avg_corr,
                    '課程結構類型': structure_type,
                    '教育特徵': interpretation,
                    '有效課程對數': len(college_correlations)
                })
                
                self.update_results(f"  平均相關性: {avg_corr:.3f} ({structure_type})")
                self.update_results(f"  結構特徵: {interpretation}")
        
        # 學院間比較分析
        if len(college_results) >= 2:
            self.update_results("\n【學院間相關性比較】")
            
            # 找出相關性最高和最低的學院
            max_college = max(college_results, key=lambda x: x['平均相關性'])
            min_college = min(college_results, key=lambda x: x['平均相關性'])
            
            self.update_results(f"課程整合度最高: {max_college['學院']} (平均r={max_college['平均相關性']:.3f})")
            self.update_results(f"課程整合度最低: {min_college['學院']} (平均r={min_college['平均相關性']:.3f})")
            
            # 分析不同課程結構類型的分布
            structure_types = {}
            for college in college_results:
                struct_type = college['課程結構類型']
                if struct_type not in structure_types:
                    structure_types[struct_type] = []
                structure_types[struct_type].append(college['學院'])
            
            self.update_results("\n課程結構類型分布:")
            for struct_type, colleges_list in structure_types.items():
                self.update_results(f"  {struct_type}: {', '.join(colleges_list)}")
            
            # 特殊發現提示
            correlations_range = max_college['平均相關性'] - min_college['平均相關性']
            if correlations_range > 0.3:
                self.update_results(f"\n【重要發現】學院間課程整合度差異顯著 (差異={correlations_range:.3f})")
                self.update_results("建議深入研究不同學院的課程設計理念差異")
            elif correlations_range > 0.15:
                self.update_results(f"\n學院間存在中等程度的課程結構差異 (差異={correlations_range:.3f})")
            else:
                self.update_results(f"\n各學院課程結構相對一致 (差異={correlations_range:.3f})")
        
        return college_results, detailed_correlations

def main():
    root = tk.Tk()
    app = CorrelationAnalysisGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()