import os
import time
from win32com.client import Dispatch

class ExcelColumnClearer:
    def __init__(self):
        self.folder_path = ""
        self.column_number = 0

    def set_folder_path(self, folder_path):
        self.folder_path = folder_path

    def set_column_number(self, column_number):
        self.column_number = column_number

    def clear_column_in_files(self):
        if not self.folder_path or self.column_number <= 0:
            raise ValueError("请先设置有效的文件夹路径和列号")

        processed_files = 0
        excel_app = None
        
        # 计算总文件数
        total_files = 0
        for root, dirs, files in os.walk(self.folder_path):
            for file in files:
                if file.endswith(('.xlsx', '.xls')):
                    total_files += 1
        
        print(f"共找到 {total_files} 个Excel文件")
        start_time = time.time()

        try:
            # 创建Excel应用实例
            excel_app = Dispatch('Excel.Application')
            excel_app.Visible = False
            excel_app.DisplayAlerts = False

            # 遍历目标文件夹中的所有Excel文件
            for root, dirs, files in os.walk(self.folder_path):
                for file in files:
                    if file.endswith(('.xlsx', '.xls')):
                        file_path = os.path.join(root, file)
                        processed_files += 1
                        # 显示进度
                        print(f"\r正在处理: {processed_files}/{total_files} - {file}", end="")
                        try:
                            # 使用COM接口打开工作簿
                            wb = excel_app.Workbooks.Open(file_path)
                            ws = wb.ActiveSheet

                            # 清空指定列的内容（跳过表头）
                            last_row = ws.UsedRange.Rows.Count
                            column = ws.Columns(self.column_number)
                            
                            # 获取列的范围（从第2行到最后一行）
                            clear_range = ws.Range(
                                ws.Cells(2, self.column_number),
                                ws.Cells(last_row, self.column_number)
                            )
                            
                            # 清空内容
                            clear_range.ClearContents()
                            
                            # 保存并关闭工作簿
                            wb.Save()
                            wb.Close()

                        except Exception as e:
                            print(f"处理文件 {file} 时出错：{str(e)}")
                            if 'wb' in locals():
                                try:
                                    wb.Close(False)
                                except:
                                    pass
                            continue

        finally:
            # 确保Excel实例被正确关闭
            if excel_app is not None:
                try:
                    excel_app.Quit()
                except:
                    pass

        # 完成后显示总结
        total_time = time.time() - start_time
        print(f"\n处理完成! 共处理 {processed_files} 个文件，耗时: {total_time:.2f}秒")
        return processed_files