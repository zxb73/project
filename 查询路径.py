import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import platform

class PathSelector:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("文件/文件夹路径选择器")
        self.root.geometry("900x700")
        
        # 设置窗口置顶
        self.root.attributes('-topmost', True)
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        title_label = tk.Label(main_frame, text="文件/文件夹路径选择工具", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 系统信息
        sys_info = f"系统类型: {platform.system()} {platform.release()}"
        sys_label = tk.Label(main_frame, text=sys_info, font=("Arial", 10))
        sys_label.pack(pady=5)
        
        # 按钮框架
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        # 选择文件按钮
        file_btn = tk.Button(button_frame, text="选择文件", 
                           command=self.select_file,
                           bg="#4CAF50", fg="white", 
                           font=("Arial", 12), width=12, height=1)
        file_btn.pack(side=tk.LEFT, padx=5)
        
        # 选择文件夹按钮
        folder_btn = tk.Button(button_frame, text="选择文件夹", 
                             command=self.select_folder,
                             bg="#2196F3", fg="white", 
                             font=("Arial", 12), width=12, height=1)
        folder_btn.pack(side=tk.LEFT, padx=5)
        
        # 选择多个文件按钮
        multi_file_btn = tk.Button(button_frame, text="选择多个文件", 
                                 command=self.select_multiple_files,
                                 bg="#FF9800", fg="white", 
                                 font=("Arial", 12), width=12, height=1)
        multi_file_btn.pack(side=tk.LEFT, padx=5)

        # 一键复制按钮框架
        copy_buttons_frame = tk.Frame(main_frame)
        copy_buttons_frame.pack(pady=10)
        
        # 复制所有路径按钮
        copy_all_btn = tk.Button(copy_buttons_frame, text="📋 复制所有路径", 
                               command=self.copy_all_paths,
                               bg="#9C27B0", fg="white", 
                               font=("Arial", 10), width=15, height=1)
        copy_all_btn.pack(side=tk.LEFT, padx=3)
        
        # 复制最后路径按钮
        copy_last_btn = tk.Button(copy_buttons_frame, text="📋 复制最后路径", 
                                command=self.copy_last_path,
                                bg="#E91E63", fg="white", 
                                font=("Arial", 10), width=15, height=1)
        copy_last_btn.pack(side=tk.LEFT, padx=3)
        
        # 复制文件名按钮
        copy_names_btn = tk.Button(copy_buttons_frame, text="📋 复制文件名", 
                                 command=self.copy_filenames,
                                 bg="#673AB7", fg="white", 
                                 font=("Arial", 10), width=15, height=1)
        copy_names_btn.pack(side=tk.LEFT, padx=3)

        # 复制父目录按钮
        copy_parent_btn = tk.Button(copy_buttons_frame, text="📋 复制父目录", 
                                  command=self.copy_parent_dirs,
                                  bg="#009688", fg="white", 
                                  font=("Arial", 10), width=15, height=1)
        copy_parent_btn.pack(side=tk.LEFT, padx=3)
        
        # 结果显示区域
        result_frame = tk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 结果标签和快速操作
        result_header_frame = tk.Frame(result_frame)
        result_header_frame.pack(fill=tk.X, pady=5)
        
        result_label = tk.Label(result_header_frame, text="选择的路径:", 
                               font=("Arial", 12, "bold"))
        result_label.pack(side=tk.LEFT)
        
        # 快速操作按钮
        quick_actions_frame = tk.Frame(result_header_frame)
        quick_actions_frame.pack(side=tk.RIGHT)
        
        # 快速复制选中文本按钮
        self.copy_selected_btn = tk.Button(quick_actions_frame, text="复制选中文本", 
                                         command=self.copy_selected_text,
                                         bg="#FF5722", fg="white",
                                         font=("Arial", 8), width=12)
        self.copy_selected_btn.pack(side=tk.LEFT, padx=2)
        
        # 滚动文本框显示路径
        self.result_text = scrolledtext.ScrolledText(result_frame, 
                                                    height=15,
                                                    font=("Consolas", 10),
                                                    wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 操作按钮框架
        action_frame = tk.Frame(main_frame)
        action_frame.pack(pady=10)
        
        # 清空按钮
        clear_btn = tk.Button(action_frame, text="清空所有", 
                            command=self.clear_results,
                            bg="#f44336", fg="white",
                            font=("Arial", 10), width=12)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 导出到文件按钮
        export_btn = tk.Button(action_frame, text="导出到文件", 
                             command=self.export_to_file,
                             bg="#607D8B", fg="white",
                             font=("Arial", 10), width=12)
        export_btn.pack(side=tk.LEFT, padx=5)
        
        # 统计信息标签
        self.stats_label = tk.Label(main_frame, text="当前选择: 0 个路径", 
                                   font=("Arial", 10), fg="gray")
        self.stats_label.pack(pady=5)
        
        # 状态栏
        self.status_label = tk.Label(main_frame, text="就绪", 
                                    font=("Arial", 9), fg="blue",
                                    relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X, pady=5)
        
    def select_file(self):
        """选择单个文件"""
        self.update_status("正在选择文件...")
        file_path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[("所有文件", "*.*")]
        )
        
        if file_path:
            absolute_path = os.path.abspath(file_path)
            self.display_path(absolute_path, "文件")
            self.update_status(f"已添加文件: {os.path.basename(file_path)}")
        else:
            self.update_status("文件选择已取消")
            
    def select_folder(self):
        """选择文件夹"""
        self.update_status("正在选择文件夹...")
        folder_path = filedialog.askdirectory(title="选择文件夹")
        
        if folder_path:
            absolute_path = os.path.abspath(folder_path)
            self.display_path(absolute_path, "文件夹")
            self.update_status(f"已添加文件夹: {os.path.basename(folder_path)}")
        else:
            self.update_status("文件夹选择已取消")
            
    def select_multiple_files(self):
        """选择多个文件"""
        self.update_status("正在选择多个文件...")
        file_paths = filedialog.askopenfilenames(
            title="选择多个文件",
            filetypes=[("所有文件", "*.*")]
        )
        
        if file_paths:
            for file_path in file_paths:
                absolute_path = os.path.abspath(file_path)
                self.display_path(absolute_path, "文件")
            self.update_status(f"已添加 {len(file_paths)} 个文件")
        else:
            self.update_status("多文件选择已取消")
    
    def display_path(self, path, path_type):
        """显示路径到文本框"""
        current_content = self.result_text.get(1.0, tk.END).strip()
        
        if current_content:
            new_content = f"{current_content}\n{path}"
        else:
            new_content = path
        
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, new_content)
        
        # 更新统计信息
        self.update_stats()
        
        # 在控制台也输出路径（可选）
        print(f"{path_type}路径: {path}")
    
    def clear_results(self):
        """清空所有结果"""
        self.result_text.delete(1.0, tk.END)
        self.update_stats()
        self.update_status("已清空所有路径")
        messagebox.showinfo("提示", "已清空所有路径")
    
    def copy_all_paths(self):
        """复制所有路径到剪贴板"""
        paths = self.result_text.get(1.0, tk.END).strip()
        if paths:
            self.root.clipboard_clear()
            self.root.clipboard_append(paths)
            self.update_status(f"已复制 {len(paths.splitlines())} 个路径到剪贴板")
            messagebox.showinfo("成功", f"所有路径已复制到剪贴板\n共 {len(paths.splitlines())} 个路径")
        else:
            self.update_status("复制失败: 没有路径可复制")
            messagebox.showwarning("警告", "没有路径可复制")
    
    def copy_last_path(self):
        """复制最后一条路径"""
        paths = self.result_text.get(1.0, tk.END).strip()
        if paths:
            last_path = paths.splitlines()[-1]
            self.root.clipboard_clear()
            self.root.clipboard_append(last_path)
            self.update_status("已复制最后一条路径到剪贴板")
            messagebox.showinfo("成功", f"最后一条路径已复制:\n{last_path}")
        else:
            self.update_status("复制失败: 没有路径可复制")
            messagebox.showwarning("警告", "没有路径可复制")
    
    def copy_filenames(self):
        """复制所有文件名（不含路径）"""
        paths = self.result_text.get(1.0, tk.END).strip()
        if paths:
            filenames = []
            for path in paths.splitlines():
                if os.path.exists(path):
                    filenames.append(os.path.basename(path))
            
            if filenames:
                filename_text = "\n".join(filenames)
                self.root.clipboard_clear()
                self.root.clipboard_append(filename_text)
                self.update_status(f"已复制 {len(filenames)} 个文件名到剪贴板")
                messagebox.showinfo("成功", f"所有文件名已复制到剪贴板\n共 {len(filenames)} 个文件")
            else:
                self.update_status("复制失败: 无法获取文件名")
                messagebox.showwarning("警告", "无法获取文件名")
        else:
            self.update_status("复制失败: 没有路径可复制")
            messagebox.showwarning("警告", "没有路径可复制")
    
    def copy_parent_dirs(self):
        """复制所有父目录路径"""
        paths = self.result_text.get(1.0, tk.END).strip()
        if paths:
            parent_dirs = []
            for path in paths.splitlines():
                if os.path.exists(path):
                    parent_dirs.append(os.path.dirname(path))
            
            if parent_dirs:
                # 去重
                unique_dirs = list(dict.fromkeys(parent_dirs))
                dirs_text = "\n".join(unique_dirs)
                self.root.clipboard_clear()
                self.root.clipboard_append(dirs_text)
                self.update_status(f"已复制 {len(unique_dirs)} 个父目录到剪贴板")
                messagebox.showinfo("成功", f"所有父目录已复制到剪贴板\n共 {len(unique_dirs)} 个目录")
            else:
                self.update_status("复制失败: 无法获取父目录")
                messagebox.showwarning("警告", "无法获取父目录")
        else:
            self.update_status("复制失败: 没有路径可复制")
            messagebox.showwarning("警告", "没有路径可复制")
    
    def copy_selected_text(self):
        """复制选中的文本"""
        try:
            selected_text = self.result_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text.strip():
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.update_status("已复制选中文本到剪贴板")
            else:
                self.update_status("没有选中的文本")
        except tk.TclError:
            self.update_status("请先选择要复制的文本")
            messagebox.showwarning("警告", "请先选择要复制的文本")
    
    def export_to_file(self):
        """导出路径到文件"""
        paths = self.result_text.get(1.0, tk.END).strip()
        if not paths:
            messagebox.showwarning("警告", "没有路径可导出")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="导出路径到文件",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(paths)
                self.update_status(f"路径已导出到: {file_path}")
                messagebox.showinfo("成功", f"路径已成功导出到:\n{file_path}")
            except Exception as e:
                self.update_status(f"导出失败: {str(e)}")
                messagebox.showerror("错误", f"导出文件时出错:\n{str(e)}")
    
    def update_stats(self):
        """更新统计信息"""
        content = self.result_text.get(1.0, tk.END).strip()
        if content:
            path_count = len(content.split('\n'))
            self.stats_label.config(text=f"当前选择: {path_count} 个路径")
        else:
            self.stats_label.config(text="当前选择: 0 个路径")
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
        print(f"状态: {message}")
    
    def run(self):
        """运行应用"""
        self.update_status("就绪 - 请选择文件或文件夹")
        self.root.mainloop()

def main():
    """主函数"""
    print("启动文件/文件夹路径选择器...")
    print("功能说明:")
    print("  • 选择文件/文件夹获取绝对路径")
    print("  • 支持多种一键复制功能")
    print("  • 支持导出路径到文件")
    print("-" * 50)
    
    app = PathSelector()
    app.run()

if __name__ == "__main__":
    main()