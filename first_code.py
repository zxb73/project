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

class DeepSeekExcelAnalyzerGUI:
    def __init__(self):
        # 在这里写死API密钥（请替换为你的实际密钥）
        self.API_KEY = "sk-2df6ea0568774004950cd5eb2e2adc8a"  # 请替换为你的真实API密钥
        
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("DeepSeek Excel 智能分析工具 - 股票分析版")
        self.root.geometry("800x700")
        self.root.configure(bg='#f0f0f0')
        
        # 初始化变量
        self.file_path = tk.StringVar()
        self.analysis_type = tk.StringVar(value="general")
        self.custom_prompt = tk.StringVar()
        self.is_analyzing = False
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        """创建所有界面组件"""
        # 创建标题
        title_font = tkFont.Font(family="Microsoft YaHei", size=16, weight="bold")
        title_label = tk.Label(
            self.root, 
            text="📊 DeepSeek Excel 股票分析工具", 
            font=title_font, 
            bg='#f0f0f0', 
            fg='#2c3e50'
        )
        title_label.pack(pady=20)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 步骤1: 文件选择
        self.create_file_section(main_frame)
        
        # 步骤2: 文件信息显示
        self.create_file_info_section(main_frame)
        
        # 步骤3: 分析选项
        self.create_analysis_section(main_frame)
        
        # 步骤4: 自定义分析需求
        self.create_custom_prompt_section(main_frame)
        
        # 步骤5: 控制按钮
        self.create_control_buttons(main_frame)
        
        # 步骤6: 进度和结果显示
        self.create_progress_section(main_frame)
        
        # 步骤7: 日志显示
        self.create_log_section(main_frame)
        
        # 显示API状态
        self.create_api_status_section(main_frame)
    
    def create_api_status_section(self, parent):
        """显示API密钥状态"""
        status_frame = ttk.LabelFrame(parent, text="🔑 API状态", padding=10)
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
    
    def create_file_section(self, parent):
        """创建文件选择区域"""
        file_frame = ttk.LabelFrame(parent, text="📁 步骤1: 选择Excel文件", padding=10)
        file_frame.pack(fill=tk.X, pady=5)
        
        # 文件路径显示和选择按钮
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X)
        
        tk.Label(path_frame, text="文件路径:").pack(side=tk.LEFT, padx=5)
        
        path_entry = ttk.Entry(path_frame, textvariable=self.file_path, width=50)
        path_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(
            path_frame, 
            text="浏览文件", 
            command=self.browse_file
        ).pack(side=tk.RIGHT, padx=5)
    
    def create_file_info_section(self, parent):
        """创建文件信息显示区域"""
        self.file_info_frame = ttk.LabelFrame(parent, text="📊 文件信息", padding=10)
        self.file_info_frame.pack(fill=tk.X, pady=5)
        
        # 初始提示文本
        self.file_info_text = tk.Text(
            self.file_info_frame, 
            height=4, 
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.file_info_text.pack(fill=tk.X)
        
        # 设置初始提示
        self.update_file_info("请先选择Excel文件")
    
    def create_analysis_section(self, parent):
        """创建分析选项区域"""
        analysis_frame = ttk.LabelFrame(parent, text="🔍 步骤2: 选择分析类型", padding=10)
        analysis_frame.pack(fill=tk.X, pady=5)
        
        # 分析类型选择 - 增加股票分析选项
        analysis_types = [
            ("股票技术分析", "stock_technical"),
            ("股票基本面分析", "stock_fundamental"),
            ("股票趋势分析", "stock_trend"),
            ("常规数据分析", "general"),
            ("财务数据分析", "finance"),
            ("市场趋势分析", "market")
        ]
        
        # 创建两行布局
        for i, (text, value) in enumerate(analysis_types):
            row = i // 3
            col = i % 3
            ttk.Radiobutton(
                analysis_frame, 
                text=text, 
                variable=self.analysis_type, 
                value=value
            ).grid(row=row, column=col, sticky="w", padx=5, pady=2)
    
    def create_custom_prompt_section(self, parent):
        """创建自定义分析需求区域"""
        custom_frame = ttk.LabelFrame(parent, text="💡 步骤3: 自定义分析需求（可选）", padding=10)
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
        example_text = "例如：分析MACD指标、RSI超买超卖情况、支撑阻力位、成交量分析等"
        self.custom_text.insert("1.0", example_text)
    
    def create_control_buttons(self, parent):
        """创建控制按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=15)
        
        self.analyze_button = ttk.Button(
            button_frame,
            text="开始分析",
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
    
    def create_progress_section(self, parent):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(parent, text="⏳ 分析进度", padding=10)
        progress_frame.pack(fill=tk.X, pady=5)
        
        # 进度变量和进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            mode='determinate'  # 明确模式
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
        
        # 状态标签
        self.status_label = tk.Label(
            progress_frame, 
            text="等待开始分析...",
            bg='#f0f0f0',
            font=("Microsoft YaHei", 9)
        )
        self.status_label.pack(pady=5)
    
    def create_log_section(self, parent):
        """创建日志显示区域"""
        log_frame = ttk.LabelFrame(parent, text="📝 分析日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=8, 
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def browse_file(self):
        """浏览文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            self.file_path.set(file_path)
            self.load_file_info(file_path)
    
    def detect_file_type(self, file_path):
        """检测文件类型并返回合适的读取参数"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.xlsx':
            # 新格式Excel文件
            return {
                'engines': ['openpyxl', 'xlrd'],
                'encodings': ['utf-8', 'gbk', 'latin1']
            }
        elif file_ext == '.xls':
            # 旧格式Excel文件
            return {
                'engines': ['xlrd', 'openpyxl'],
                'encodings': ['gbk', 'utf-8', 'latin1']
            }
        else:
            # 未知格式，尝试所有可能
            return {
                'engines': ['xlrd', 'openpyxl'],
                'encodings': ['gbk', 'utf-8', 'latin1']
            }
    
    def read_excel_with_encoding(self, file_path):
        """使用多种方式读取Excel文件，处理编码问题"""
        file_type_info = self.detect_file_type(file_path)
        engines = file_type_info['engines']
        encodings = file_type_info['encodings']
        
        attempts = []
        
        # 生成所有尝试组合
        for engine in engines:
            for encoding in encodings:
                attempts.append({'engine': engine, 'encoding': encoding})
            # 也尝试不指定编码
            attempts.append({'engine': engine})
        
        # 添加更多特殊尝试
        special_attempts = [
            {'engine': None},  # 让pandas自动选择
            {'engine': 'xlrd', 'encoding': 'gbk', 'na_values': ['', ' ', 'NULL', 'null']},
            {'engine': 'openpyxl', 'encoding': 'gbk', 'na_values': ['', ' ', 'NULL', 'null']},
        ]
        
        attempts.extend(special_attempts)
        
        for i, kwargs in enumerate(attempts):
            try:
                self.log_message(f"🔄 尝试读取方式 {i+1}: {kwargs}")
                
                # 移除encoding参数如果引擎不支持
                if 'encoding' in kwargs and kwargs['engine'] == 'openpyxl':
                    # openpyxl不支持encoding参数
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
        
        # 如果所有方式都失败，尝试使用错误处理方式
        try:
            self.log_message("🔄 尝试最终读取方式: 忽略所有错误")
            # 尝试读取为CSV再转换（作为最后的手段）
            if file_path.endswith('.xls'):
                # 对于.xls文件，尝试使用xlrd的特殊参数
                try:
                    import xlrd
                    df = pd.read_excel(file_path, engine='xlrd', encoding_override='gbk')
                    if not df.empty:
                        fixed_columns = self.fix_column_names(df.columns.tolist())
                        df.columns = fixed_columns
                        return df
                except:
                    pass
            
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
                
        except Exception as e:
            self.log_message(f"❌ 最终读取方式也失败: {str(e)}")
        
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
    
    def load_file_info(self, file_path):
        """加载并显示文件信息"""
        try:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                self.update_file_info("错误: 文件不存在")
                self.log_message("❌ 文件不存在")
                return
            
            # 检查文件大小
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                self.update_file_info("错误: 文件为空")
                self.log_message("❌ 文件为空")
                return
            
            # 使用改进的读取方法
            df = self.read_excel_with_encoding(file_path)
            if df is None:
                self.update_file_info("错误: 无法读取Excel文件，请检查文件格式和内容")
                self.log_message("❌ 所有读取方式都失败")
                return
            
            file_size_kb = file_size / 1024  # KB
            
            # 检查是否包含股票数据常见列
            stock_columns = ['开盘', '收盘', '最高', '最低', '成交量', '涨跌幅', 'open', 'close', 'high', 'low', 'volume']
            has_stock_data = any(any(stock_col in str(col).lower() for stock_col in stock_columns) for col in df.columns)
            
            # 显示列名信息（限制显示数量）
            display_columns = df.columns.tolist()[:10]  # 只显示前10列
            columns_display = ', '.join([str(col) for col in display_columns])
            if len(df.columns) > 10:
                columns_display += f' ... (共{len(df.columns)}列)'
            
            info_text = f"""
文件名称: {os.path.basename(file_path)}
文件大小: {file_size_kb:.1f} KB
数据规模: {df.shape[0]} 行 × {df.shape[1]} 列
文件格式: {os.path.splitext(file_path)[1].upper()}
列名: {columns_display}

数据类型统计:
- 数值型: {len(df.select_dtypes(include=['number']).columns)} 列
- 文本型: {len(df.select_dtypes(include=['object']).columns)} 列
- 日期型: {len(df.select_dtypes(include=['datetime']).columns)} 列
- 股票数据: {'✅ 检测到股票数据' if has_stock_data else '⚠️ 未检测到标准股票列名'}
            """
            
            self.update_file_info(info_text.strip())
            self.log_message(f"✅ 成功加载文件: {os.path.basename(file_path)}")
            self.log_message(f"📊 数据形状: {df.shape[0]} 行 × {df.shape[1]} 列")
            
        except Exception as e:
            self.update_file_info(f"错误: 无法读取文件\n{str(e)}")
            self.log_message(f"❌ 文件读取错误: {str(e)}")
    
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
    
    def update_progress(self, value, message):
        """更新进度条和状态"""
        self.progress_var.set(value)
        self.progress_percent.config(text=f"{int(value)}%")
        self.status_label.config(text=message)
        
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
            
            "general": "请对这个数据集进行全面的数据分析，包括数据质量评估、关键指标识别、趋势分析和业务建议",
            "finance": "请分析财务数据，包括收入、成本、利润等关键财务指标，评估财务健康状况",
            "market": "请分析市场数据，包括市场份额、竞争分析、市场趋势预测"
        }
        
        base_prompt = base_prompts.get(self.analysis_type.get(), base_prompts["general"])
        custom_text = self.custom_text.get(1.0, tk.END).strip()
        
        if custom_text and custom_text != "例如：分析MACD指标、RSI超买超卖情况、支撑阻力位、成交量分析等":
            return f"{base_prompt}。特别关注：{custom_text}"
        else:
            return base_prompt
    
    def start_analysis(self):
        """开始分析（在新线程中运行）"""
        if self.is_analyzing:
            return
        
        # 验证输入
        if not self.file_path.get():
            messagebox.showerror("错误", "请选择Excel文件")
            return
        
        # 验证API密钥
        if not self.API_KEY or self.API_KEY == "sk-your-api-key-here":
            messagebox.showerror("错误", "请先配置API密钥")
            return
        
        # 验证文件是否存在
        if not os.path.exists(self.file_path.get()):
            messagebox.showerror("错误", "选择的文件不存在")
            return
        
        # 禁用按钮，开始分析
        self.is_analyzing = True
        self.analyze_button.config(state=tk.DISABLED)
        self.log_message("🚀 开始分析过程...")
        
        # 重置进度条
        self.update_progress(0, "初始化分析环境...")
        
        # 在新线程中运行分析，避免界面冻结
        analysis_thread = threading.Thread(target=self.run_analysis)
        analysis_thread.daemon = True
        analysis_thread.start()
    
    def run_analysis(self):
        """执行分析过程"""
        try:
            # 步骤1: 读取数据
            self.update_progress(10, "读取Excel文件...")
            self.log_message("步骤1: 读取Excel文件")
            
            data_info = self.read_excel_file(self.file_path.get())
            if not data_info:
                self.analysis_complete(False, "读取Excel文件失败")
                return
            
            # 短暂暂停让进度条可见
            threading.Event().wait(0.5)
            
            # 步骤2: 准备分析
            self.update_progress(30, "准备分析数据...")
            self.log_message("步骤2: 准备分析数据")
            analysis_prompt = self.get_analysis_prompt()
            
            threading.Event().wait(0.5)
            
            # 步骤3: 调用DeepSeek API
            self.update_progress(50, "调用DeepSeek API进行分析...")
            self.log_message("步骤3: 使用DeepSeek AI进行分析")
            
            analysis_result = self.analyze_with_deepseek(data_info, analysis_prompt)
            if not analysis_result:
                self.analysis_complete(False, "DeepSeek分析失败")
                return
            
            threading.Event().wait(0.5)
            
            # 步骤4: 保存结果
            self.update_progress(80, "保存分析结果...")
            self.log_message("步骤4: 保存分析结果到桌面")
            
            saved_path = self.save_results(analysis_result, data_info, analysis_prompt)
            if not saved_path:
                self.analysis_complete(False, "保存结果失败")
                return
            
            threading.Event().wait(0.5)
            
            # 步骤5: 完成
            self.update_progress(100, "分析完成！")
            self.log_message("✅ 分析完成！")
            
            self.analysis_complete(True, f"分析完成！结果已保存到:\n{saved_path}")
            
        except Exception as e:
            self.analysis_complete(False, f"分析过程中出错: {str(e)}")
    
    def read_excel_file(self, file_path):
        """读取Excel文件"""
        try:
            df = self.read_excel_with_encoding(file_path)
            if df is None:
                return None
                
            return {
                'dataframe': df,
                'shape': df.shape,
                'columns': df.columns.tolist(),
                'dtypes': df.dtypes.to_dict(),
                'null_counts': df.isnull().sum().to_dict(),
                'file_path': file_path,
                'file_name': os.path.basename(file_path),
                'file_size': os.path.getsize(file_path) / 1024
            }
        except Exception as e:
            self.log_message(f"❌ 读取文件错误: {str(e)}")
            return None
    
    def analyze_with_deepseek(self, data_info, analysis_request):
        """使用DeepSeek API分析数据"""
        try:
            client = OpenAI(
                api_key=self.API_KEY,  # 使用写死的API密钥
                base_url="https://api.deepseek.com"
            )
            
            data_summary = f"""
数据集基本信息:
- 文件名: {data_info['file_name']}
- 数据形状: {data_info['shape']}
- 列名: {', '.join([str(col) for col in data_info['columns']])}
- 数据类型: {data_info['dtypes']}
- 空值统计: {data_info['null_counts']}

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
            self.file_path.set("")
            self.analysis_type.set("general")
            self.custom_text.delete(1.0, tk.END)
            self.custom_text.insert(1.0, "例如：分析MACD指标、RSI超买超卖情况、支撑阻力位、成交量分析等")
            self.update_file_info("请先选择Excel文件")
            self.update_progress(0, "等待开始分析...")
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