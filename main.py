import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'
import tkinter as tk
from tkinter import filedialog, messagebox
from excel_processor import ExcelProcessor

class ExcelUpdaterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 批量更新工具")
        self.root.geometry("400x300")
        
        # 设置窗口背景色
        self.root.configure(bg='#f0f0f0')

        self.master_file_path = ""
        self.target_folder = ""
        self.processor = ExcelProcessor(print)

        # 添加匹配列和更新列选择
        self.match_column_var = tk.StringVar(value="2")
        self.match_column_options = ["2", "3"]
        
        self.update_column_var = tk.StringVar(value="3")
        self.update_column_options = ["3", "4"]

        self.create_widgets()

    def create_widgets(self):
        # 统一按钮样式
        button_style = {
            'bg': '#4a90e2',  # 按钮背景色
            'fg': 'white',    # 按钮文字颜色
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
        btn_master = tk.Button(self.root, text="选择 Master 总表", **button_style, command=self.select_master_file)
        btn_master.pack(pady=10)
        self.master_label = tk.Label(self.root, text="未选择文件", **label_style)
        self.master_label.pack()

        btn_folder = tk.Button(self.root, text="选择目标文件夹", **button_style, command=self.select_target_folder)
        btn_folder.pack(pady=10)
        self.folder_label = tk.Label(self.root, text="未选择文件夹", **label_style)
        self.folder_label.pack()

        # 匹配列选择
        match_frame = tk.Frame(self.root, bg='#f0f0f0')
        match_frame.pack(pady=10)
        tk.Label(match_frame, text="匹配列：", **label_style).pack(side=tk.LEFT)
        match_dropdown = tk.OptionMenu(match_frame, self.match_column_var, *self.match_column_options)
        match_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        match_dropdown["menu"].config(bg='white', fg='#333333')
        match_dropdown.pack(side=tk.LEFT)
        tk.Label(match_frame, text="列", **label_style).pack(side=tk.LEFT)

        # 更新列选择
        update_frame = tk.Frame(self.root, bg='#f0f0f0')
        update_frame.pack(pady=10)
        tk.Label(update_frame, text="更新列：", **label_style).pack(side=tk.LEFT)
        update_dropdown = tk.OptionMenu(update_frame, self.update_column_var, *self.update_column_options)
        update_dropdown.config(bg='#4a90e2', fg='white', font=('Arial', 10), width=5)
        update_dropdown["menu"].config(bg='white', fg='#333333')
        update_dropdown.pack(side=tk.LEFT)
        tk.Label(update_frame, text="列", **label_style).pack(side=tk.LEFT)

        # 执行按钮
        btn_start = tk.Button(self.root, text="开始处理", **button_style, command=self.process_files)
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
            update_column = int(self.update_column_var.get()) - 1
            if match_column < 0 or update_column < 0:
                raise ValueError("列索引必须大于0")
            # 设置匹配列和更新列索引
            self.processor.set_match_column(match_column)
            self.processor.set_update_column(update_column)
        except ValueError as e:
            messagebox.showerror("错误", f"匹配列设置错误：{str(e)}")
            return

        try:
            updated_count = self.processor.process_files()
            messagebox.showinfo("完成", f"共更新 {updated_count} 行。")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelUpdaterGUI()
    app.run()