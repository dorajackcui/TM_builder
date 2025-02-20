import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'
import tkinter as tk
from tkinter import filedialog, messagebox
from excel_processor import ExcelProcessor

class ExcelUpdaterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 批量更新工具")
        self.root.geometry("300x200")

        self.master_file_path = ""
        self.target_folder = ""
        self.processor = ExcelProcessor(print)

        self.create_widgets()

    def create_widgets(self):
        # 文件选择按钮
        btn_master = tk.Button(self.root, text="选择 Master 总表", command=self.select_master_file)
        btn_master.pack(pady=10)
        self.master_label = tk.Label(self.root, text="未选择文件")
        self.master_label.pack()

        btn_folder = tk.Button(self.root, text="选择目标文件夹", command=self.select_target_folder)
        btn_folder.pack(pady=10)
        self.folder_label = tk.Label(self.root, text="未选择文件夹")
        self.folder_label.pack()

        # 执行按钮
        btn_start = tk.Button(self.root, text="开始处理", command=self.process_files)
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
            updated_count = self.processor.process_files()
            messagebox.showinfo("完成", f"共更新 {updated_count} 个单元格。")
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelUpdaterGUI()
    app.run()