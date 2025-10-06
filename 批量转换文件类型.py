import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def select_file():
    """选择文件并返回文件路径"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.attributes('-topmost', True)  # 窗口置顶
    
    file_path = filedialog.askopenfilename(
        title="请选择 Excel 文件",
        filetypes=[
            ("Excel files", "*.xls *.xlsx *.xlsm"),
            ("All files", "*.*")
        ]
    )
    
    root.destroy()
    return file_path

def read_excel_file(file_path):
    """
    尝试多种方式读取 Excel 文件
    """
    if not file_path:
        return None
        
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    print(f"正在读取文件: {file_path}")
    print(f"文件大小: {os.path.getsize(file_path)} 字节")
    
    # 方法1：尝试用 openpyxl 引擎（用于 .xlsx）
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print("✓ 使用 openpyxl 引擎读取成功")
        return df
    except Exception as e:
        print(f"✗ openpyxl 读取失败: {e}")
    
    # 方法2：尝试用 xlrd 引擎（用于 .xls）
    try:
        df = pd.read_excel(file_path, engine='xlrd')
        print("✓ 使用 xlrd 引擎读取成功")
        return df
    except Exception as e:
        print(f"✗ xlrd 读取失败: {e}")
    
    # 方法3：不指定引擎，让 pandas 自动选择
    try:
        df = pd.read_excel(file_path)
        print("✓ 自动选择引擎读取成功")
        return df
    except Exception as e:
        print(f"✗ 自动选择引擎读取失败: {e}")
    
    # 方法4：尝试用 calamine 引擎
    try:
        df = pd.read_excel(file_path, engine='calamine')
        print("✓ 使用 calamine 引擎读取成功")
        return df
    except Exception as e:
        print(f"✗ calamine 读取失败: {e}")
    
    raise Exception("所有读取方法都失败了")

def show_data_preview(df, file_path):
    """显示数据预览"""
    root = tk.Tk()
    root.title(f"数据预览 - {os.path.basename(file_path)}")
    root.geometry("800x600")
    
    # 创建文本框显示数据
    text_widget = tk.Text(root, wrap=tk.NONE)
    text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # 添加滚动条
    scrollbar_y = tk.Scrollbar(text_widget, orient=tk.VERTICAL, command=text_widget.yview)
    scrollbar_x = tk.Scrollbar(text_widget, orient=tk.HORIZONTAL, command=text_widget.xview)
    text_widget.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    # 显示数据基本信息
    info_text = f"文件: {os.path.basename(file_path)}\n"
    info_text += f"数据形状: {df.shape} (行: {df.shape[0]}, 列: {df.shape[1]})\n"
    info_text += f"列名: {list(df.columns)}\n"
    info_text += "-" * 50 + "\n"
    info_text += "前10行数据:\n\n"
    
    # 显示前10行数据
    preview_data = df.head(10).to_string()
    info_text += preview_data
    
    text_widget.insert(tk.END, info_text)
    text_widget.config(state=tk.DISABLED)  # 设置为只读
    
    # 添加保存按钮
    def save_successful():
        messagebox.showinfo("成功", "数据读取成功！")
        root.destroy()
    
    tk.Button(root, text="确认", command=save_successful, bg="green", fg="white").pack(pady=10)
    
    root.mainloop()

def main():
    """主函数"""
    print("=== Excel 文件读取工具 ===")
    
    # 选择文件
    file_path = select_file()
    
    if not file_path:
        print("未选择文件，程序退出")
        return
    
    try:
        # 读取文件
        data = read_excel_file(file_path)
        
        if data is not None:
            print(f"\n✓ 读取成功！")
            print(f"数据形状: {data.shape}")
            print(f"列名: {list(data.columns)}")
            print(f"\n前5行数据:")
            print(data.head())
            
            # 显示图形化预览
            show_data_preview(data, file_path)
            
        else:
            messagebox.showerror("错误", "读取文件失败")
            
    except Exception as e:
        error_msg = f"读取文件时出错:\n{str(e)}"
        print(error_msg)
        messagebox.showerror("错误", error_msg)

if __name__ == "__main__":
    main()