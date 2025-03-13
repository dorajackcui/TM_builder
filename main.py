import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from excel_processor import ExcelProcessor
from excel_cleaner import ExcelColumnClearer
from excel_compatibility_processor import ExcelCompatibilityProcessor
from multi_column_processor import MultiColumnExcelProcessor

class ExcelUpdaterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 工具集")
        self.root.geometry("400x500")
        
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
        
        # 创建兼容性处理选项卡
        self.compatibility_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.compatibility_frame, text='兼容性处理')
        
        # 创建多列更新选项卡
        self.multi_column_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.multi_column_frame, text='多列更新')
        
        # 初始化所有工具
        self.init_updater()
        self.init_clearer()
        self.init_compatibility()
        self.init_multi_column()

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
        tk.Label(column_frame, text="列号：", **label_style).pack(side=tk.LEFT)
        self.column_var = tk.StringVar(value="")
        column_entry = tk.Entry(column_frame, textvariable=self.column_var, width=5)
        column_entry.pack(side=tk.LEFT)
        tk.Label(column_frame, text="列", **label_style).pack(side=tk.LEFT)

        # 功能按钮
        btn_clear = tk.Button(self.clearer_frame, text="清空列", **button_style, command=self.clear_column)
        btn_clear.pack(pady=5)
        
        btn_insert = tk.Button(self.clearer_frame, text="插入列", **button_style, command=self.insert_column)
        btn_insert.pack(pady=5)
        
        btn_delete = tk.Button(self.clearer_frame, text="删除列", **button_style, command=self.delete_column)
        btn_delete.pack(pady=5)

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

    def insert_column(self):
        try:
            column_number = int(self.column_var.get())
            if column_number <= 0:
                raise ValueError("列号必须大于0")
            self.clearer.set_column_number(column_number)
            processed_files = self.clearer.insert_column_in_files()
            messagebox.showinfo("完成", f"共处理 {processed_files} 个文件。")
        except ValueError as e:
            messagebox.showerror("错误", f"列号设置错误：{str(e)}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def delete_column(self):
        try:
            column_number = int(self.column_var.get())
            if column_number <= 0:
                raise ValueError("列号必须大于0")
            self.clearer.set_column_number(column_number)

            # 弹出确认对话框
            confirm = messagebox.askyesno("确认操作", f"确定要删除所有Excel文件的第{column_number}列吗？\n此操作不可撤销！")
            if not confirm:
                return

            processed_files = self.clearer.delete_column_in_files()
            messagebox.showinfo("完成", f"共处理 {processed_files} 个文件。")
        except ValueError as e:
            messagebox.showerror("错误", f"列号设置错误：{str(e)}")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def init_compatibility(self):
        self.compatibility_processor = ExcelCompatibilityProcessor()

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
        btn_folder = tk.Button(self.compatibility_frame, text="选择目标文件夹", **button_style, command=self.select_compatibility_folder)
        btn_folder.pack(pady=10)
        self.compatibility_folder_label = tk.Label(self.compatibility_frame, text="未选择文件夹", **label_style)
        self.compatibility_folder_label.pack()

        # 执行按钮
        btn_start = tk.Button(self.compatibility_frame, text="开始处理", **button_style, command=self.process_compatibility)
        btn_start.pack(pady=10)

    def select_compatibility_folder(self):
        folder_path = filedialog.askdirectory(title="选择目标文件夹")
        if folder_path:
            self.compatibility_folder_label.config(text=f"已选择：{os.path.basename(folder_path)}")
            self.compatibility_processor.set_folder_path(folder_path)

    def process_compatibility(self):
        try:
            processed_files = self.compatibility_processor.process_files()
            messagebox.showinfo("完成", f"共处理 {processed_files} 个文件。")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            
    def init_multi_column(self):
        self.multi_processor = MultiColumnExcelProcessor(print)

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
        btn_master = tk.Button(self.multi_column_frame, text="选择 Master 总表", **button_style, command=self.select_multi_master_file)
        btn_master.pack(pady=10)
        self.multi_master_label = tk.Label(self.multi_column_frame, text="未选择文件", **label_style)
        self.multi_master_label.pack()

        btn_folder = tk.Button(self.multi_column_frame, text="选择目标文件夹", **button_style, command=self.select_multi_target_folder)
        btn_folder.pack(pady=10)
        self.multi_folder_label = tk.Label(self.multi_column_frame, text="未选择文件夹", **label_style)
        self.multi_folder_label.pack()

        # 匹配列选择
        self.multi_match_column_var = tk.StringVar(value="3")
        self.multi_match_column_options = ["4", "3"]
        match_frame = tk.Frame(self.multi_column_frame, bg='#f0f0f0')
        match_frame.pack(pady=10)
        tk.Label(match_frame, text="匹配列：", **label_style).pack(side=tk.LEFT)
        match_dropdown = tk.OptionMenu(match_frame, self.multi_match_column_var, *self.multi_match_column_options)
        match_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        match_dropdown["menu"].config(bg='white', fg='#333333')
        match_dropdown.pack(side=tk.LEFT)
        tk.Label(match_frame, text="列", **label_style).pack(side=tk.LEFT)

        # 开始内容列选择（从master表获取数据的起始列）
        self.multi_start_column_var = tk.StringVar(value="5")
        self.multi_start_column_options = ["4","5"]
        start_frame = tk.Frame(self.multi_column_frame, bg='#f0f0f0')
        start_frame.pack(pady=10)
        tk.Label(start_frame, text="开始内容列：", **label_style).pack(side=tk.LEFT)
        start_dropdown = tk.OptionMenu(start_frame, self.multi_start_column_var, *self.multi_start_column_options)
        start_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        start_dropdown["menu"].config(bg='white', fg='#333333')
        start_dropdown.pack(side=tk.LEFT)
        tk.Label(start_frame, text="列（Master表）", **label_style).pack(side=tk.LEFT)

        # 更新开始列选择（目标文件要更新的起始列）
        self.multi_update_start_column_var = tk.StringVar(value="5")
        self.multi_update_start_column_options = ["4", "5"]
        update_frame = tk.Frame(self.multi_column_frame, bg='#f0f0f0')
        update_frame.pack(pady=10)
        tk.Label(update_frame, text="更新开始列：", **label_style).pack(side=tk.LEFT)
        update_dropdown = tk.OptionMenu(update_frame, self.multi_update_start_column_var, *self.multi_update_start_column_options)
        update_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        update_dropdown["menu"].config(bg='white', fg='#333333')
        update_dropdown.pack(side=tk.LEFT)
        tk.Label(update_frame, text="列（目标文件）", **label_style).pack(side=tk.LEFT)
        
        # 列数选择
        self.multi_column_count_var = tk.StringVar(value="7")
        self.multi_column_count_options = ["8", "7", "6", "5", "4"]
        count_frame = tk.Frame(self.multi_column_frame, bg='#f0f0f0')
        count_frame.pack(pady=10)
        tk.Label(count_frame, text="更新列数：", **label_style).pack(side=tk.LEFT)
        count_dropdown = tk.OptionMenu(count_frame, self.multi_column_count_var, *self.multi_column_count_options)
        count_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        count_dropdown["menu"].config(bg='white', fg='#333333')
        count_dropdown.pack(side=tk.LEFT)
        tk.Label(count_frame, text="列", **label_style).pack(side=tk.LEFT)

        # 执行按钮
        btn_start = tk.Button(self.multi_column_frame, text="开始处理", **button_style, command=self.process_multi_column)
        btn_start.pack(pady=10)
        
    def select_multi_master_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Master 总表",
            filetypes=[("Excel 文件", "*.xlsx *.xls")]
        )
        if file_path:
            self.multi_master_label.config(text=f"已选择：{os.path.basename(file_path)}")
            self.multi_processor.set_master_file(file_path)

    def select_multi_target_folder(self):
        folder_path = filedialog.askdirectory(title="选择目标文件夹")
        if folder_path:
            self.multi_folder_label.config(text=f"已选择：{os.path.basename(folder_path)}")
            self.multi_processor.set_target_folder(folder_path)
            
    def process_multi_column(self):
        try:
            # 将下拉菜单选择的值转换为0基索引
            match_column = int(self.multi_match_column_var.get()) - 1
            start_column = int(self.multi_start_column_var.get()) - 1
            update_start_column = int(self.multi_update_start_column_var.get()) - 1
            column_count = int(self.multi_column_count_var.get())
            
            if match_column < 0 or start_column < 0 or update_start_column < 0 or column_count <= 0:
                raise ValueError("列索引必须大于0，列数必须大于0")
                
            # 设置匹配列、开始列、更新开始列和列数
            self.multi_processor.set_match_column(match_column)
            self.multi_processor.set_start_column(start_column)
            self.multi_processor.set_update_start_column(update_start_column)
            self.multi_processor.set_column_count(column_count)
            
            updated_count = self.multi_processor.process_files()
            messagebox.showinfo("完成", f"共更新 {updated_count} 处数据。")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelUpdaterGUI()
    app.run()