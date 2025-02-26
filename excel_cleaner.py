import os
import win32com.client

class ExcelColumnClearer:
    def __init__(self):
        self.folder_path = ""
        self.column_number = 0
        self.excel = None

    def set_folder_path(self, folder_path):
        self.folder_path = folder_path

    def set_column_number(self, column_number):
        self.column_number = column_number

    def clear_column_in_files(self):
        if not self.folder_path or self.column_number <= 0:
            raise ValueError("请先设置有效的文件夹路径和列号")

        processed_files = 0
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False

        try:
            for root, dirs, files in os.walk(self.folder_path):
                for file in files:
                    if file.endswith(('.xlsx', '.xls')):
                        file_path = os.path.join(root, file)
                        try:
                            # 使用pywin32打开工作簿
                            wb = self.excel.Workbooks.Open(file_path)
                            ws = wb.ActiveSheet
                            
                            # 获取最后一行
                            last_row = ws.UsedRange.Rows.Count
                            
                            # 清空指定列的内容（跳过表头）
                            if last_row > 1:
                                clear_range = ws.Range(
                                    ws.Cells(2, self.column_number),
                                    ws.Cells(last_row, self.column_number)
                                )
                                clear_range.ClearContents()
                            
                            # 保存并关闭工作簿
                            wb.Save()
                            wb.Close()
                            processed_files += 1

                        except Exception as e:
                            print(f"处理文件 {file} 时出错：{str(e)}")
                            if 'wb' in locals():
                                wb.Close(SaveChanges=False)
                            continue

        finally:
            if self.excel:
                self.excel.Quit()
                self.excel = None

        return processed_files