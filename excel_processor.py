import os
import concurrent.futures
import win32com.client

class ExcelProcessor:
    def __init__(self, log_callback=None):
        self.master_file_path = ""
        self.target_folder = ""
        self.log_callback = log_callback or (lambda msg: None)
        self.master_columns = []  # 存储列位置信息
        self.match_column_index = 1  # 默认使用第二列作为匹配列
        self.update_column_index = 2  # 默认更新第三列
        self.excel = None

    def set_update_column(self, column_index):
        """设置要更新的列索引"""
        self.update_column_index = column_index

    def set_match_column(self, column_index):
        """设置用于匹配的列索引"""
        self.match_column_index = column_index

    def set_master_file(self, file_path):
        self.master_file_path = file_path

    def set_target_folder(self, folder_path):
        self.target_folder = folder_path

    def log(self, message):
        self.log_callback(message)

    def process_files(self):
        if not self.master_file_path or not self.target_folder:
            raise ValueError("请先选择 Master 文件和目标文件夹！")

        try:
            self.log("正在读取 Master 文件...")
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

            # 读取Master文件
            wb = self.excel.Workbooks.Open(self.master_file_path)
            ws = wb.ActiveSheet
            last_row = ws.UsedRange.Rows.Count

            # 创建master_dict，只读取必要的列
            master_dict = {}
            key_column = 2  # B列作为key
            # Master表：匹配列和更新列需要+1，因为它们是相对于GUI选择的列号
            master_match_column = self.match_column_index + 2  # GUI选择的匹配列+1
            master_update_column = self.update_column_index + 2  # GUI选择的更新列+1
            # 批量读取所需列的数据
            key_range = ws.Range(ws.Cells(2, key_column), ws.Cells(last_row, key_column))
            match_range = ws.Range(ws.Cells(2, master_match_column), ws.Cells(last_row, master_match_column))
            update_range = ws.Range(ws.Cells(2, master_update_column), ws.Cells(last_row, master_update_column))

            # 一次性获取所有值
            key_values = key_range.Value
            match_values = match_range.Value
            update_values = update_range.Value

            # 处理数据并创建master_dict
            for i in range(len(key_values)):
                try:
                    # 更安全的值获取和类型转换
                    key_value = None
                    match_value = None
                    update_value = None

                    # 使用异常处理来安全地获取值
                    try:
                        key_value = key_values[i][0] if key_values[i] else None
                    except:
                        continue

                    try:
                        match_value = match_values[i][0] if match_values[i] else None
                    except:
                        continue

                    try:
                        update_value = update_values[i][0] if update_values[i] else None
                    except:
                        continue
                    
                    # 确保值不是COM错误对象并进行类型检查
                    if key_value is not None and not isinstance(key_value, Exception):
                        if isinstance(key_value, (int, float, str)):
                            key = str(key_value).strip()
                            if key and match_value is not None and update_value is not None:
                                if not isinstance(match_value, Exception) and not isinstance(update_value, Exception):
                                    if isinstance(match_value, (int, float, str)) and isinstance(update_value, (int, float, str)):
                                        match_val = str(match_value)
                                        update_val = str(update_value)
                                        if match_val:
                                            master_dict[key] = [match_val, update_val]
                except Exception as e:
                    self.log(f"处理第{i+2}行数据时出错：{str(e)}")
                    continue

            wb.Close(SaveChanges=False)
            self.log(f"Master 中共找到 {len(master_dict)} 个有效 Key")

            # 收集目标文件
            file_paths = []
            for root, _, files in os.walk(self.target_folder):
                file_paths.extend(
                    os.path.join(root, file)
                    for file in files
                    if file.lower().endswith(('.xlsx', '.xls'))
                )
            self.log(f"找到 {len(file_paths)} 个目标文件")

            # 优化：减少线程池大小，避免过多的COM对象创建
            max_workers = min(4, len(file_paths))  # 限制最大线程数为4
            updated_count = 0
            
            # 分批处理文件
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 每次提交一批文件进行处理
                batch_size = max(1, len(file_paths) // max_workers)
                for i in range(0, len(file_paths), batch_size):
                    batch = file_paths[i:i + batch_size]
                    futures = [executor.submit(self._process_single_file, fp, master_dict) for fp in batch]
                    # 等待当前批次完成
                    for future in concurrent.futures.as_completed(futures):
                        try:
                            result = future.result()
                            updated_count += result
                        except Exception as e:
                            self.log(f"处理文件批次时出错：{str(e)}")

            self.log(f"处理完成，共更新 {updated_count} 处数据")
            return updated_count
        finally:
            if self.excel:
                try:
                    self.excel.Quit()
                except:
                    pass
                self.excel = None

    def _process_single_file(self, file_path, master_dict):
        excel = None
        wb = None
        updated = 0

        try:
            # 创建新的Excel实例
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            # 打开工作簿
            wb = excel.Workbooks.Open(file_path)
            ws = wb.ActiveSheet
            last_row = ws.UsedRange.Rows.Count

            updates = {}
            # 读取和处理数据
            for row in range(2, last_row + 1):
                try:
                    target_key = str(ws.Cells(row, 1).Value or '').strip()  # A列
                    target_match_value = str(ws.Cells(row, self.match_column_index + 1).Value or '')  # 目标表使用GUI选择的列号

                    if not target_key or not target_match_value:
                        continue

                    if target_key in master_dict:
                        master_values = master_dict[target_key]
                        if target_match_value == master_values[0]:
                            updates[row] = master_values[1]
                            updated += 1
                except Exception as e:
                    self.log(f"处理文件 {file_path} 第 {row} 行时出错：{str(e)}")
                    continue

            # 批量更新单元格
            if updates:
                try:
                    for row, value in updates.items():
                        ws.Cells(row, self.update_column_index + 1).Value = value
                    wb.Save()
                except Exception as e:
                    self.log(f"保存文件 {file_path} 时出错：{str(e)}")
                    return 0

            return updated
        except Exception as e:
            self.log(f"处理文件 {file_path} 时出错：{str(e)}")
            return 0
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
                if excel is not None:
                    excel.Quit()
            except:
                pass