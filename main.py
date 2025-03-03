import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from excel_processor import ExcelProcessor
from excel_cleaner import ExcelColumnClearer

class ExcelUpdaterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 工具集")
        self.root.geometry("400x400")
        
        # 设置窗口背景色
        self.root.configure(bg='#f0f0f0')

        self.master_file_path = ""
        self.target_folder = ""
        self.processor = ExcelProcessor(print)

        # 添加匹配列、内容列和更新列选择
        self.match_column_var = tk.StringVar(value="2")
        self.match_column_options = ["2", "3"]
        
        self.content_column_var = tk.StringVar(value="4")
        self.content_column_options = ["4","5","6","7","8","9","10","11"]
        
        self.update_column_var = tk.StringVar(value="3")
        self.update_column_options = ["3", "4"]
        style = ttk.Style()
        style.theme_use('clam')  # 使用兼容性更好的主题
        style.configure('TNotebook', background='#f0f0f0', borderwidth=0)
        style.configure('TNotebook.Tab', 
            padding=[20, 8], 
            background='#e0e0e0', 
            foreground='#333333', 
            borderwidth=1,
            font=('Arial', 10)
        )
        style.map('TNotebook.Tab',
            padding=[('selected', [20, 8])],
            background=[('selected', '#4a90e2'), ('active', '#b8d6f5')],
            foreground=[('selected', 'white'), ('active', '#333333')]
        )
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TNotebook.Tab', 
            padding=[20, 8], 
            background='#e0e0e0', 
            foreground='#333333', 
            borderwidth=1,
            font=('Arial', 10)
        )
        style.map('TNotebook.Tab',
            background=[('selected', '#4a90e2'), ('active', '#b8d6f5')],
            foreground=[('selected', 'white'), ('active', '#333333')]
        )
        style.configure('TFrame', background='#f0f0f0')

        # 创建选项卡控件
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)

        # 创建批量更新工具选项卡
        self.updater_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.updater_frame, text='批量更新')

        # 创建列清空工具选项卡
        self.clearer_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.clearer_frame, text='列清空')

        # 初始化两个工具
        self.init_updater()
        self.init_clearer()

    def init_updater(self):
        # 统一按钮样式
        button_style = {
            'bg': '#4a90e2',
            'fg': 'white',
            'font': ('Arial', 10),
            'relief': 'raised',
            'padx': 20
        }
        
        # 统一标签样式
        label_style = {
            'bg': '#f0f0f0',  # 标签背景色
            'fg': '#333333',  # 标签文字颜色
            'font': ('Arial', 10)
        }

        # 文件选择按钮
        btn_master = tk.Button(self.updater_frame, text="选择 Master 总表", **button_style, command=self.select_master_file)
        btn_master.pack(pady=10)
        self.master_label = tk.Label(self.updater_frame, text="未选择文件", **label_style)
        self.master_label.pack()

        btn_folder = tk.Button(self.updater_frame, text="选择目标文件夹", **button_style, command=self.select_target_folder)
        btn_folder.pack(pady=10)
        self.folder_label = tk.Label(self.updater_frame, text="未选择文件夹", **label_style)
        self.folder_label.pack()

        # 匹配列选择
        match_frame = tk.Frame(self.updater_frame, bg='#f0f0f0')
        match_frame.pack(pady=10)
        tk.Label(match_frame, text="匹配列：", **label_style).pack(side=tk.LEFT)
        match_dropdown = tk.OptionMenu(match_frame, self.match_column_var, *self.match_column_options)
        match_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        match_dropdown["menu"].config(bg='white', fg='#333333')
        match_dropdown.pack(side=tk.LEFT)
        tk.Label(match_frame, text="列", **label_style).pack(side=tk.LEFT)

        # 内容列选择（从master表获取数据的列）
        content_frame = tk.Frame(self.updater_frame, bg='#f0f0f0')
        content_frame.pack(pady=10)
        tk.Label(content_frame, text="内容列：", **label_style).pack(side=tk.LEFT)
        content_dropdown = tk.OptionMenu(content_frame, self.content_column_var, *self.content_column_options)
        content_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        content_dropdown["menu"].config(bg='white', fg='#333333')
        content_dropdown.pack(side=tk.LEFT)
        tk.Label(content_frame, text="列（Master表）", **label_style).pack(side=tk.LEFT)

        # 更新列选择（目标文件要更新的列）
        update_frame = tk.Frame(self.updater_frame, bg='#f0f0f0')
        update_frame.pack(pady=10)
        tk.Label(update_frame, text="更新列：", **label_style).pack(side=tk.LEFT)
        update_dropdown = tk.OptionMenu(update_frame, self.update_column_var, *self.update_column_options)
        update_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        update_dropdown["menu"].config(bg='white', fg='#333333')
        update_dropdown.pack(side=tk.LEFT)
        tk.Label(update_frame, text="列（目标文件）", **label_style).pack(side=tk.LEFT)

        # 执行按钮
        btn_start = tk.Button(self.updater_frame, text="开始处理", **button_style, command=self.process_files)
        btn_start.pack(pady=10)

    def select_master_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Master 总表",
            filetypes=[("Excel 文件", "*.xlsx *.xls")]
        )
        if file_path:
            self.master_file_path = file_path
            self.master_label.config(text=f"已选择：{os.path.basename(file_path)}")
            self.processor.set_master_file(file_path)

    def select_target_folder(self):
        folder_path = filedialog.askdirectory(title="选择目标文件夹")
        if folder_path:
            self.target_folder = folder_path
            self.folder_label.config(text=f"已选择：{os.path.basename(folder_path)}")
            self.processor.set_target_folder(folder_path)

    def process_files(self):
        if not self.master_file_path or not self.target_folder:
            messagebox.showerror("错误", "请先选择 Master 文件和目标文件夹！")
            return

        try:
            # 将下拉菜单选择的值转换为0基索引
            match_column = int(self.match_column_var.get()) - 1
            content_column = int(self.content_column_var.get()) - 1
            update_column = int(self.update_column_var.get()) - 1
            if match_column < 0 or content_column < 0 or update_column < 0:
                raise ValueError("列索引必须大于0")
            # 设置匹配列、内容列和更新列索引
            self.processor.set_match_column(match_column)
            self.processor.set_content_column(content_column)
            self.processor.set_update_column(update_column)
        except ValueError as e:
            messagebox.showerror("错误", f"匹配列设置错误：{str(e)}")
            return

        try:
            updated_count = self.processor.process_files()
            messagebox.showinfo("完成", f"共更新 {updated_count} 行。")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def init_clearer(self):
        self.clearer = ExcelColumnClearer()

        # 统一按钮样式
        button_style = {
            'bg': '#4a90e2',
            'fg': 'white',
            'font': ('Arial', 10),
            'relief': 'raised',
            'padx': 20
        }
        
        # 统一标签样式
        label_style = {
            'bg': '#f0f0f0',  # 标签背景色
            'fg': '#333333',  # 标签文字颜色
            'font': ('Arial', 10)
        }

        # 文件夹选择按钮
        btn_folder = tk.Button(self.clearer_frame, text="选择目标文件夹", **button_style, command=self.select_clearer_folder)
        btn_folder.pack(pady=10)
        self.clearer_folder_label = tk.Label(self.clearer_frame, text="未选择文件夹", **label_style)
        self.clearer_folder_label.pack()

        # 列号输入框
        column_frame = tk.Frame(self.clearer_frame, bg='#f0f0f0')
        column_frame.pack(pady=10)
        tk.Label(column_frame, text="清空列号：", **label_style).pack(side=tk.LEFT)
        self.column_var = tk.StringVar(value="")
        column_entry = tk.Entry(column_frame, textvariable=self.column_var, width=5)
        column_entry.pack(side=tk.LEFT)
        tk.Label(column_frame, text="列", **label_style).pack(side=tk.LEFT)

        # 执行按钮
        btn_start = tk.Button(self.clearer_frame, text="开始清空", **button_style, command=self.clear_column)
        btn_start.pack(pady=10)

    def select_clearer_folder(self):
        folder_path = filedialog.askdirectory(title="选择目标文件夹")
        if folder_path:
            self.clearer_folder_label.config(text=f"已选择：{os.path.basename(folder_path)}")
            self.clearer.set_folder_path(folder_path)

    def clear_column(self):
        try:
            column_number = int(self.column_var.get())
            if column_number <= 0:
                raise ValueError("列号必须大于0")
            self.clearer.set_column_number(column_number)
            processed_files = self.clearer.clear_column_in_files()
            messagebox.showinfo("完成", f"共处理 {processed_files} 个文件。")
        except ValueError as e:
            messagebox.showerror("错误", f"列号设置错误：{str(e)}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelUpdaterGUI()
    app.run()