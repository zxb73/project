import os
import pandas as pd
import json
from openai import OpenAI
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
from tkinter import font as tkFont
import sys
import chardet
import glob
import re
import xlwings as xw

class DeepSeekExcelAnalyzerGUI:
    def __init__(self):
        # 在这里写死API密钥（请替换为你的实际密钥）
        self.API_KEY = "sk-2df6ea0568774004950cd5eb2e2adc8a"  # 请替换为你的真实API密钥
        
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("DeepSeek Excel 智能分析工具 - 股票分析版")
        self.root.geometry("900x600")  # 设置初始窗口大小
        self.root.configure(bg='#f0f0f0')
        
        # 初始化变量
        self.folder_path = tk.StringVar()
        self.analysis_type = tk.StringVar(value="stock_technical")
        self.custom_prompt = tk.StringVar()
        self.is_analyzing = False
        self.selected_files = []  # 存储选择的文件列表
        self.processed_files = 0  # 已处理文件计数
        
        # 创建带滚动条的界面
        self.create_scrollable_interface()
        
    def create_scrollable_interface(self):
        """创建带滚动条的界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建Canvas和滚动条
        self.canvas = tk.Canvas(main_frame, bg='#f0f0f0')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # 绑定鼠标滚轮事件
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)
        
        # 布局Canvas和滚动条
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 创建界面组件
        self.create_widgets()
        
    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def create_widgets(self):
        """创建所有界面组件"""
        # 创建标题
        title_font = tkFont.Font(family="Microsoft YaHei", size=16, weight="bold")
        title_label = tk.Label(
            self.scrollable_frame, 
            text="📊 DeepSeek Excel 股票分析工具 - 批量版", 
            font=title_font, 
            bg='#f0f0f0', 
            fg='#2c3e50'
        )
        title_label.pack(pady=20)
        
        # 显示API状态
        self.create_api_status_section()
        
        # 步骤1: 文件夹选择
        self.create_folder_section()
        
        # 步骤2: 文件列表显示
        self.create_file_list_section()
        
        # 步骤3: 文件信息显示
        self.create_file_info_section()
        
        # 步骤4: 分析选项
        self.create_analysis_section()
        
        # 步骤5: 自定义分析需求
        self.create_custom_prompt_section()
        
        # 步骤6: 进度和结果显示
        self.create_progress_section()
        
        # 步骤7: 控制按钮
        self.create_control_buttons()
        
        # 步骤8: 日志显示
        self.create_log_section()
    
    def create_api_status_section(self):
        """显示API密钥状态"""
        status_frame = ttk.LabelFrame(self.scrollable_frame, text="🔑 API状态", padding=10)
        status_frame.pack(fill=tk.X, pady=5)
        
        # 显示API密钥状态（隐藏部分字符）
        masked_key = self.API_KEY[:10] + "***" + self.API_KEY[-4:] if len(self.API_KEY) > 14 else "***"
        status_text = f"API密钥已配置: {masked_key}"
        
        status_label = tk.Label(
            status_frame, 
            text=status_text,
            bg='#e8f5e8',
            fg='#2e7d32',
            font=("Microsoft YaHei", 10)
        )
        status_label.pack(fill=tk.X, pady=5)
    
    def create_folder_section(self):
        """创建文件夹选择区域"""
        folder_frame = ttk.LabelFrame(self.scrollable_frame, text="📁 步骤1: 选择数据文件夹", padding=10)
        folder_frame.pack(fill=tk.X, pady=5)
        
        # 文件夹路径显示和选择按钮
        path_frame = ttk.Frame(folder_frame)
        path_frame.pack(fill=tk.X)
        
        tk.Label(path_frame, text="文件夹路径:").pack(side=tk.LEFT, padx=5)
        
        path_entry = ttk.Entry(path_frame, textvariable=self.folder_path, width=50)
        path_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(
            path_frame, 
            text="浏览文件夹", 
            command=self.browse_folder
        ).pack(side=tk.RIGHT, padx=5)
        
        # 文件筛选选项
        filter_frame = ttk.Frame(folder_frame)
        filter_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(filter_frame, text="文件类型:").pack(side=tk.LEFT, padx=5)
        
        self.file_pattern = tk.StringVar(value="*.xlsx")
        ttk.Radiobutton(filter_frame, text="Excel文件(*.xlsx)", variable=self.file_pattern, value="*.xlsx").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(filter_frame, text="Excel文件(*.xls)", variable=self.file_pattern, value="*.xls").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(filter_frame, text="所有Excel文件", variable=self.file_pattern, value="*.xls*").pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            filter_frame,
            text="扫描文件",
            command=self.scan_files
        ).pack(side=tk.RIGHT, padx=5)
    
    def create_file_list_section(self):
        """创建文件列表显示区域"""
        self.file_list_frame = ttk.LabelFrame(self.scrollable_frame, text="📋 步骤2: 文件列表", padding=10)
        self.file_list_frame.pack(fill=tk.X, pady=5)
        
        # 文件列表和选择框
        list_frame = ttk.Frame(self.file_list_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧文件列表
        left_frame = ttk.Frame(list_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        tk.Label(left_frame, text="检测到的文件:").pack(anchor="w")
        
        # 文件列表框（固定高度）
        listbox_frame = ttk.Frame(left_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.file_listbox = tk.Listbox(
            listbox_frame, 
            selectmode=tk.MULTIPLE,
            height=6
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 文件列表滚动条
        listbox_scrollbar = ttk.Scrollbar(listbox_frame, command=self.file_listbox.yview)
        listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=listbox_scrollbar.set)
        
        # 右侧操作按钮
        right_frame = ttk.Frame(list_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        
        ttk.Button(
            right_frame,
            text="全选",
            command=self.select_all_files,
            width=12
        ).pack(pady=2)
        
        ttk.Button(
            right_frame,
            text="全不选",
            command=self.deselect_all_files,
            width=12
        ).pack(pady=2)
        
        ttk.Button(
            right_frame,
            text="反选",
            command=self.invert_selection,
            width=12
        ).pack(pady=2)
        
        # 文件计数
        self.file_count_label = tk.Label(
            self.file_list_frame,
            text="共检测到 0 个文件",
            font=("Microsoft YaHei", 9)
        )
        self.file_count_label.pack(anchor="w")
    
    def create_file_info_section(self):
        """创建文件信息显示区域"""
        self.file_info_frame = ttk.LabelFrame(self.scrollable_frame, text="📊 文件信息预览", padding=10)
        self.file_info_frame.pack(fill=tk.X, pady=5)
        
        # 初始提示文本（固定高度）
        self.file_info_text = scrolledtext.ScrolledText(
            self.file_info_frame, 
            height=4, 
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Microsoft YaHei", 9)
        )
        self.file_info_text.pack(fill=tk.X)
        
        # 设置初始提示
        self.update_file_info("请先选择文件夹并扫描文件")
    
    def create_analysis_section(self):
        """创建分析选项区域"""
        analysis_frame = ttk.LabelFrame(self.scrollable_frame, text="🔍 步骤3: 选择分析类型", padding=10)
        analysis_frame.pack(fill=tk.X, pady=5)
        
        # 分析类型选择 - 增加股票分析选项
        analysis_types = [
            ("股票技术分析", "stock_technical"),
            ("股票基本面分析", "stock_fundamental"),
            ("股票趋势分析", "stock_trend"),
            ("批量对比分析", "batch_comparison"),
            ("常规数据分析", "general")
        ]
        
        # 创建布局
        for i, (text, value) in enumerate(analysis_types):
            ttk.Radiobutton(
                analysis_frame, 
                text=text, 
                variable=self.analysis_type, 
                value=value
            ).grid(row=i//3, column=i%3, sticky="w", padx=5, pady=2)
    
    def create_custom_prompt_section(self):
        """创建自定义分析需求区域"""
        custom_frame = ttk.LabelFrame(self.scrollable_frame, text="💡 步骤4: 自定义分析需求（可选）", padding=10)
        custom_frame.pack(fill=tk.X, pady=5)
        
        prompt_label = tk.Label(
            custom_frame, 
            text="如有特殊分析需求，请在此输入:",
            wraplength=700
        )
        prompt_label.pack(anchor="w", pady=(0, 5))
        
        self.custom_text = scrolledtext.ScrolledText(
            custom_frame, 
            height=4, 
            wrap=tk.WORD,
            font=("Microsoft YaHei", 10)
        )
        self.custom_text.pack(fill=tk.X)
        
        # 股票分析示例提示
        example_text = "例如：分析MACD指标、RSI超买超卖情况、支撑阻力位、成交量分析等。对于批量分析，可以要求对比不同日期的数据趋势。"
        self.custom_text.insert("1.0", example_text)
    
    def create_progress_section(self):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(self.scrollable_frame, text="⏳ 分析进度", padding=10)
        progress_frame.pack(fill=tk.X, pady=5)
        
        # 进度变量和进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # 进度百分比标签
        self.progress_percent = tk.Label(
            progress_frame,
            text="0%",
            font=("Microsoft YaHei", 10, "bold"),
            fg="#2c3e50"
        )
        self.progress_percent.pack()
        
        # 文件进度标签
        self.file_progress_label = tk.Label(
            progress_frame,
            text="等待开始分析...",
            font=("Microsoft YaHei", 9),
            fg="#666666"
        )
        self.file_progress_label.pack()
        
        # 状态标签
        self.status_label = tk.Label(
            progress_frame, 
            text="等待开始分析...",
            bg='#f0f0f0',
            font=("Microsoft YaHei", 9)
        )
        self.status_label.pack(pady=5)
    
    def create_control_buttons(self):
        """创建控制按钮区域"""
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.pack(fill=tk.X, pady=15)
        
        self.analyze_button = ttk.Button(
            button_frame,
            text="开始批量分析",
            command=self.start_analysis,
            style="Accent.TButton"
        )
        self.analyze_button.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(
            button_frame,
            text="清空重来",
            command=self.reset_all
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(
            button_frame,
            text="退出程序",
            command=self.root.quit
        ).pack(side=tk.RIGHT, padx=10)
    
    def create_log_section(self):
        """创建日志显示区域"""
        log_frame = ttk.LabelFrame(self.scrollable_frame, text="📝 分析日志", padding=10)
        log_frame.pack(fill=tk.X, pady=5)
        
        # 固定高度的日志显示区域
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=8, 
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.X)
    
    def browse_folder(self):
        """浏览文件夹"""
        folder_path = filedialog.askdirectory(title="选择数据文件夹")
        
        if folder_path:
            self.folder_path.set(folder_path)
            self.log_message(f"✅ 已选择文件夹: {folder_path}")
    
    def scan_files(self):
        """扫描文件夹中的Excel文件"""
        folder_path = self.folder_path.get()
        if not folder_path or not os.path.exists(folder_path):
            messagebox.showerror("错误", "请先选择有效的文件夹路径")
            return
        
        try:
            # 清空文件列表
            self.file_listbox.delete(0, tk.END)
            self.selected_files = []
            
            # 根据选择的文件类型扫描文件
            pattern = self.file_pattern.get()
            search_pattern = os.path.join(folder_path, "**", pattern) if pattern == "*.xls*" else os.path.join(folder_path, pattern)
            
            # 递归搜索文件
            files = []
            if "**" in search_pattern:
                # 递归搜索子文件夹
                for file_path in glob.glob(search_pattern, recursive=True):
                    if os.path.isfile(file_path):
                        files.append(file_path)
            else:
                # 只在当前文件夹搜索
                for file_path in glob.glob(search_pattern):
                    if os.path.isfile(file_path):
                        files.append(file_path)
            
            # 按文件名排序
            files.sort()
            
            # 添加到列表
            for file_path in files:
                file_name = os.path.basename(file_path)
                relative_path = os.path.relpath(file_path, folder_path)
                display_text = f"{file_name} ({relative_path})"
                self.file_listbox.insert(tk.END, display_text)
                self.selected_files.append(file_path)
            
            # 更新文件计数
            file_count = len(files)
            self.file_count_label.config(text=f"共检测到 {file_count} 个文件")
            
            if file_count > 0:
                # 默认选择所有文件
                self.select_all_files()
                self.log_message(f"✅ 扫描完成，找到 {file_count} 个Excel文件")
                self.update_file_info(f"已扫描到 {file_count} 个文件。请选择要分析的文件，然后点击'开始批量分析'。")
            else:
                self.log_message("⚠️ 未找到匹配的Excel文件")
                self.update_file_info("未找到匹配的Excel文件，请检查文件夹路径和文件类型设置。")
                
        except Exception as e:
            self.log_message(f"❌ 扫描文件时出错: {str(e)}")
            messagebox.showerror("错误", f"扫描文件时出错: {str(e)}")
    
    def select_all_files(self):
        """选择所有文件"""
        self.file_listbox.select_set(0, tk.END)
    
    def deselect_all_files(self):
        """取消选择所有文件"""
        self.file_listbox.select_clear(0, tk.END)
    
    def invert_selection(self):
        """反选文件"""
        for i in range(self.file_listbox.size()):
            if self.file_listbox.selection_includes(i):
                self.file_listbox.selection_clear(i)
            else:
                self.file_listbox.select_set(i)
    
    def get_selected_files(self):
        """获取选中的文件列表"""
        selected_indices = self.file_listbox.curselection()
        return [self.selected_files[i] for i in selected_indices]
    
    def extract_date_from_path(self, file_path):
        """从文件路径中提取日期"""
        try:
            # 从路径中查找8位数字（YYYYMMDD格式）
            path_parts = file_path.split(os.sep)
            
            # 查找包含8位数字的文件夹名
            for part in path_parts:
                match = re.search(r'(\d{8})', part)
                if match:
                    date_str = match.group(1)
                    # 转换为YYYY-MM-DD格式
                    year = date_str[:4]
                    month = date_str[4:6]
                    day = date_str[6:8]
                    return f"{year}-{month}-{day}"
            
            # 如果没有找到，使用文件修改日期
            file_mtime = os.path.getmtime(file_path)
            return datetime.fromtimestamp(file_mtime).strftime("%Y-%m-%d")
            
        except:
            # 如果都失败，使用当前日期
            return datetime.now().strftime("%Y-%m-%d")
    
    def update_file_info(self, text):
        """更新文件信息显示"""
        self.file_info_text.config(state=tk.NORMAL)
        self.file_info_text.delete(1.0, tk.END)
        self.file_info_text.insert(1.0, text)
        self.file_info_text.config(state=tk.DISABLED)
    
    def log_message(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # 更新界面
        self.root.update_idletasks()
    
    def update_progress(self, value, message, file_progress=""):
        """更新进度条和状态"""
        self.progress_var.set(value)
        self.progress_percent.config(text=f"{int(value)}%")
        self.status_label.config(text=message)
        
        if file_progress:
            self.file_progress_label.config(text=file_progress)
        
        # 强制更新界面
        self.root.update_idletasks()
        self.root.update()
    
    def get_analysis_prompt(self):
        """获取分析提示词"""
        base_prompts = {
            "stock_technical": """请对股票数据进行深入的技术分析，包括：
- K线形态分析
- 均线系统（5日、10日、20日、60日均线）
- MACD指标分析
- RSI超买超卖情况
- 成交量与价格关系
- 支撑位和阻力位识别
- 买卖信号判断""",
            
            "stock_fundamental": """请对股票数据进行基本面分析，包括：
- 财务指标分析（市盈率、市净率等）
- 盈利能力分析
- 成长性评估
- 估值水平判断
- 行业对比分析
- 风险提示和建议""",
            
            "stock_trend": """请对股票数据进行趋势分析，包括：
- 短期、中期、长期趋势判断
- 趋势线分析
- 突破和回调识别
- 波动率分析
- 动量指标
- 未来走势预测""",
            
            "batch_comparison": """请对多个日期的股票数据进行对比分析，包括：
- 各日期数据的整体对比
- 关键指标的变化趋势
- 异常波动的识别
- 多日期连续分析
- 趋势预测和建议""",
            
            "general": "请对这个数据集进行全面的数据分析，包括数据质量评估、关键指标识别、趋势分析和业务建议"
        }
        
        base_prompt = base_prompts.get(self.analysis_type.get(), base_prompts["stock_technical"])
        custom_text = self.custom_text.get(1.0, tk.END).strip()
        
        if custom_text and custom_text != "例如：分析MACD指标、RSI超买超卖情况、支撑阻力位、成交量分析等。对于批量分析，可以要求对比不同日期的数据趋势。":
            return f"{base_prompt}。特别关注：{custom_text}"
        else:
            return base_prompt
    
    def start_analysis(self):
        """开始分析（在新线程中运行）"""
        if self.is_analyzing:
            return
        
        # 验证输入
        if not self.folder_path.get():
            messagebox.showerror("错误", "请选择数据文件夹")
            return
        
        # 验证API密钥
        if not self.API_KEY or self.API_KEY == "sk-your-api-key-here":
            messagebox.showerror("错误", "请先配置API密钥")
            return
        
        # 获取选中的文件
        selected_files = self.get_selected_files()
        if not selected_files:
            messagebox.showerror("错误", "请选择要分析的文件")
            return
        
        # 禁用按钮，开始分析
        self.is_analyzing = True
        self.analyze_button.config(state=tk.DISABLED)
        self.processed_files = 0
        self.log_message(f"🚀 开始批量分析，共 {len(selected_files)} 个文件...")
        
        # 重置进度条
        self.update_progress(0, "初始化分析环境...", "")
        
        # 在新线程中运行分析，避免界面冻结
        analysis_thread = threading.Thread(target=lambda: self.run_batch_analysis(selected_files))
        analysis_thread.daemon = True
        analysis_thread.start()
    
    def run_batch_analysis(self, file_paths):
        """执行批量分析过程"""
        total_files = len(file_paths)
        success_count = 0
        failed_files = []
        
        try:
            for i, file_path in enumerate(file_paths):
                # 更新进度
                file_progress = f"正在处理: {i+1}/{total_files} - {os.path.basename(file_path)}"
                progress_percent = (i / total_files) * 100
                self.update_progress(progress_percent, f"分析文件中...", file_progress)
                
                self.log_message(f"📊 处理文件 {i+1}/{total_files}: {os.path.basename(file_path)}")
                
                # 步骤1: 读取数据并添加统计日期
                self.update_progress(progress_percent + 5, "读取Excel文件...", file_progress)
                data_info = self.read_excel_file(file_path)
                
                if not data_info:
                    self.log_message(f"❌ 文件读取失败: {os.path.basename(file_path)}")
                    failed_files.append((file_path, "读取失败"))
                    continue
                
                # 步骤2: 准备分析
                self.update_progress(progress_percent + 10, "准备分析数据...", file_progress)
                analysis_prompt = self.get_analysis_prompt()
                
                # 步骤3: 调用DeepSeek API
                self.update_progress(progress_percent + 30, "调用DeepSeek API进行分析...", file_progress)
                analysis_result = self.analyze_with_deepseek(data_info, analysis_prompt)
                
                if not analysis_result:
                    self.log_message(f"❌ DeepSeek分析失败: {os.path.basename(file_path)}")
                    failed_files.append((file_path, "分析失败"))
                    continue
                
                # 步骤4: 保存结果
                self.update_progress(progress_percent + 60, "保存分析结果...", file_progress)
                saved_path = self.save_results(analysis_result, data_info, analysis_prompt)
                
                if not saved_path:
                    self.log_message(f"❌ 保存结果失败: {os.path.basename(file_path)}")
                    failed_files.append((file_path, "保存失败"))
                    continue
                
                success_count += 1
                self.processed_files += 1
                self.log_message(f"✅ 文件分析完成: {os.path.basename(file_path)}")
                
                # 短暂暂停
                threading.Event().wait(0.2)
            
            # 完成
            self.update_progress(100, "批量分析完成！", f"完成: {success_count}/{total_files}")
            self.log_message(f"✅ 批量分析完成！成功: {success_count}, 失败: {len(failed_files)}")
            
            if failed_files:
                failed_list = "\n".join([f"- {os.path.basename(f[0])} ({f[1]})" for f in failed_files])
                self.analysis_complete(True, f"批量分析完成！\n成功: {success_count} 个文件\n失败: {len(failed_files)} 个文件\n\n失败文件:\n{failed_list}")
            else:
                self.analysis_complete(True, f"批量分析完成！所有 {success_count} 个文件都分析成功！")
            
        except Exception as e:
            self.analysis_complete(False, f"批量分析过程中出错: {str(e)}")
    
    def read_excel_with_xlwings(self, file_path):
        """使用xlwings读取Excel文件，兼容WPS"""
        try:
            # 启动Excel应用，visible=False表示后台运行
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            sheet = wb.sheets[0]  # 获取第一个工作表
            
            # 将数据读取为pandas DataFrame
            data_range = sheet.used_range
            df = data_range.options(pd.DataFrame, index=False, header=True).value
            
            wb.close()
            app.quit()
            
            # 确保返回的是DataFrame
            if not isinstance(df, pd.DataFrame):
                # 处理读取结果不是DataFrame的情况
                if isinstance(df, list):
                    # 假设第一行是标题
                    df = pd.DataFrame(df[1:], columns=df[0])
                else:
                    # 其他情况，创建一个空的DataFrame
                    df = pd.DataFrame()
            
            return df
            
        except Exception as e:
            # 确保即使出错也尝试关闭应用
            try:
                app.quit()
            except:
                pass
            return None
    
    def read_excel_with_encoding(self, file_path):
        """使用多种方式读取Excel文件，处理编码问题"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # 尝试多种读取方式
        attempts = []
        
        # 根据文件扩展名设置不同的尝试顺序
        if file_ext == '.xlsx':
            attempts = [
                {'engine': 'openpyxl'},
                {'engine': 'openpyxl', 'encoding': 'gbk'},
                {'engine': 'openpyxl', 'encoding': 'utf-8'},
                {'engine': 'xlrd'},
                {'engine': None}  # 让pandas自动选择
            ]
        else:  # .xls文件
            attempts = [
                {'engine': 'xlrd'},
                {'engine': 'xlrd', 'encoding': 'gbk'},
                {'engine': 'xlrd', 'encoding': 'utf-8'},
                {'engine': 'openpyxl'},
                {'engine': None}  # 让pandas自动选择
            ]
        
        # 添加特殊尝试
        special_attempts = [
            {'engine': 'xlrd', 'encoding': 'gbk', 'na_values': ['', ' ', 'NULL', 'null']},
            {'engine': 'openpyxl', 'encoding': 'gbk', 'na_values': ['', ' ', 'NULL', 'null']},
        ]
        
        attempts.extend(special_attempts)
        
        for i, kwargs in enumerate(attempts):
            try:
                self.log_message(f"🔄 尝试读取方式 {i+1}: {kwargs}")
                
                # 移除encoding参数如果引擎不支持
                if 'encoding' in kwargs and kwargs['engine'] == 'openpyxl':
                    kwargs_copy = kwargs.copy()
                    del kwargs_copy['encoding']
                    df = pd.read_excel(file_path, **kwargs_copy)
                else:
                    df = pd.read_excel(file_path, **kwargs)
                
                # 修复列名
                if not df.empty:
                    original_columns = df.columns.tolist()
                    fixed_columns = self.fix_column_names(original_columns)
                    df.columns = fixed_columns
                    
                    self.log_message(f"✅ 成功读取文件，使用方式 {i+1}")
                    self.log_message(f"📋 原始列名: {original_columns}")
                    self.log_message(f"🔧 修复后列名: {fixed_columns}")
                    
                    return df
                    
            except Exception as e:
                self.log_message(f"❌ 读取方式 {i+1} 失败: {str(e)}")
                continue
        
        # 如果所有方式都失败，尝试使用xlwings
        try:
            self.log_message("🔄 尝试使用xlwings读取（WPS兼容）")
            df = self.read_excel_with_xlwings(file_path)
            if df is not None and not df.empty:
                self.log_message("✅ 使用xlwings成功读取文件")
                return df
        except Exception as e:
            self.log_message(f"❌ xlwings读取失败: {str(e)}")
        
        # 最后尝试：读取前几行来诊断问题
        try:
            df = pd.read_excel(file_path, nrows=5)
            if not df.empty:
                self.log_message("⚠️ 只能读取前5行数据，文件可能有问题")
                fixed_columns = self.fix_column_names(df.columns.tolist())
                df.columns = fixed_columns
                return df
        except:
            pass
                
        self.log_message("❌ 所有读取方式都失败")
        return None
    
    def fix_column_names(self, columns):
        """修复列名乱码问题"""
        fixed_columns = []
        
        for i, col in enumerate(columns):
            # 如果列名已经是字符串且没有乱码，直接使用
            if isinstance(col, str) and not self.has_garbled_text(col):
                fixed_columns.append(col)
                continue
            
            # 尝试不同的编码方式修复
            fixed = False
            if isinstance(col, bytes):
                encodings = ['gbk', 'utf-8', 'gb2312', 'latin1', 'cp1252']
                for encoding in encodings:
                    try:
                        decoded_col = col.decode(encoding)
                        if not self.has_garbled_text(decoded_col):
                            fixed_columns.append(decoded_col)
                            fixed = True
                            self.log_message(f"🔧 列名修复: {encoding} -> {decoded_col}")
                            break
                    except:
                        continue
            
            # 如果所有编码都失败，使用原始列名或生成新列名
            if not fixed:
                if isinstance(col, str):
                    # 尝试修复常见乱码
                    repaired_col = self.repair_garbled_text(col)
                    fixed_columns.append(repaired_col)
                    self.log_message(f"🔧 乱码修复: {col} -> {repaired_col}")
                else:
                    # 生成默认列名
                    fixed_columns.append(f"列_{i+1}")
        
        return fixed_columns
    
    def repair_garbled_text(self, text):
        """尝试修复常见的乱码文本"""
        # 常见的中文乱码映射
        garbled_map = {
            'ä¸‰': '三', 'å«': '叫', 'ç»§': '继', 'è¿': '进', 'é': '送',
            'ï¼': '，', 'ï¼': '：', 'ï¼': '！', 'ï¼': '！', 'ï¼': '＋'
        }
        
        for garbled, correct in garbled_map.items():
            text = text.replace(garbled, correct)
        
        return text
    
    def has_garbled_text(self, text):
        """检测文本是否包含乱码字符"""
        # 常见的乱码字符模式
        garbled_patterns = [
            'ä¸', 'å', 'ç»', 'è¿', 'é', 'ï¼'
        ]
        
        # 检查是否包含无法打印的字符或乱码模式
        for pattern in garbled_patterns:
            if pattern in text:
                return True
        
        # 检查是否包含大量无法识别的字符
        try:
            text.encode('utf-8')
            return False
        except:
            return True
    
    def read_excel_file(self, file_path):
        """读取Excel文件并添加统计日期字段"""
        try:
            # 使用改进的读取方法
            df = self.read_excel_with_encoding(file_path)
            if df is None:
                return None
            
            # 从文件路径中提取日期
            stat_date = self.extract_date_from_path(file_path)
            
            # 添加统计日期字段
            df['统计日期'] = stat_date
            
            return {
                'dataframe': df,
                'shape': df.shape,
                'columns': df.columns.tolist(),
                'dtypes': df.dtypes.to_dict(),
                'null_counts': df.isnull().sum().to_dict(),
                'file_path': file_path,
                'file_name': os.path.basename(file_path),
                'file_size': os.path.getsize(file_path) / 1024,
                'stat_date': stat_date  # 保存统计日期
            }
        except Exception as e:
            self.log_message(f"❌ 读取文件错误: {str(e)}")
            return None
    
    def analyze_with_deepseek(self, data_info, analysis_request):
        """使用DeepSeek API分析数据"""
        try:
            client = OpenAI(
                api_key=self.API_KEY,
                base_url="https://api.deepseek.com"
            )
            
            data_summary = f"""
数据集基本信息:
- 文件名: {data_info['file_name']}
- 数据形状: {data_info['shape']}
- 列名: {', '.join([str(col) for col in data_info['columns']])}
- 数据类型: {data_info['dtypes']}
- 空值统计: {data_info['null_counts']}
- 统计日期: {data_info['stat_date']}

分析要求: {analysis_request}
"""
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是专业的股票数据分析师，擅长技术分析和基本面分析"},
                    {"role": "user", "content": data_summary}
                ],
                stream=False,
                temperature=0.7
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            self.log_message(f"❌ API调用错误: {str(e)}")
            return None
    
    def save_results(self, analysis_result, data_info, analysis_prompt):
        """保存分析结果"""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            original_filename = data_info['file_name'].replace('.xlsx', '').replace('.xls', '')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{original_filename}_股票分析报告_{timestamp}.txt"
            output_path = os.path.join(desktop_path, output_filename)
            
            content = f"""DeepSeek 股票分析报告
生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
原始文件: {data_info['file_name']}
数据规模: {data_info['shape'][0]} 行 × {data_info['shape'][1]} 列
统计日期: {data_info['stat_date']}
分析类型: {self.analysis_type.get()}
分析需求: {analysis_prompt}

=== 数据分析结果 ===

{analysis_result}
"""
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            return output_path
            
        except Exception as e:
            self.log_message(f"❌ 保存错误: {str(e)}")
            return None
    
    def analysis_complete(self, success, message):
        """分析完成处理"""
        self.is_analyzing = False
        self.analyze_button.config(state=tk.NORMAL)
        
        if success:
            messagebox.showinfo("分析完成", message)
            self.log_message("🎉 " + message)
        else:
            messagebox.showerror("分析失败", message)
            self.log_message("❌ " + message)
    
    def reset_all(self):
        """重置所有输入"""
        if messagebox.askyesno("确认", "确定要清空所有输入并重新开始吗？"):
            self.folder_path.set("")
            self.analysis_type.set("stock_technical")
            self.custom_text.delete(1.0, tk.END)
            self.custom_text.insert(1.0, "例如：分析MACD指标、RSI超买超卖情况、支撑阻力位、成交量分析等。对于批量分析，可以要求对比不同日期的数据趋势。")
            self.update_file_info("请先选择文件夹并扫描文件")
            self.file_listbox.delete(0, tk.END)
            self.selected_files = []
            self.file_count_label.config(text="共检测到 0 个文件")
            self.update_progress(0, "等待开始分析...", "")
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
            self.log_message("系统已重置，可以开始新的分析")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

def main():
    """主函数"""
    try:
        # 创建应用程序实例
        app = DeepSeekExcelAnalyzerGUI()
        # 运行应用程序
        app.run()
    except Exception as e:
        messagebox.showerror("启动错误", f"应用程序启动失败:\n{str(e)}")

if __name__ == "__main__":
    main()