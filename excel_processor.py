import pandas as pd
import os
import concurrent.futures
import openpyxl
import time
from win32com.client import Dispatch

class ExcelProcessor:
    def __init__(self, log_callback=None):
        self.master_file_path = ""
        self.target_folder = ""
        self.log_callback = log_callback or (lambda msg: None)
        self.master_columns = []  # 存储列位置信息
        self.match_column_index = 1  # 默认使用第二列作为匹配列
        self.content_column_index = 3  # 默认使用第四列作为内容列（来自master表）
        self.update_column_index = 2  # 默认更新第三列（目标文件的列）

    def set_update_column(self, column_index):
        """设置要更新的列索引（目标文件的列）"""
        self.update_column_index = column_index

    def set_match_column(self, column_index):
        """设置用于匹配的列索引"""
        self.match_column_index = column_index
        
    def set_content_column(self, column_index):
        """设置内容列索引（master表中的列）"""
        self.content_column_index = column_index

    def set_master_file(self, file_path):
        self.master_file_path = file_path

    def set_target_folder(self, folder_path):
        self.target_folder = folder_path

    def log(self, message):
        self.log_callback(message)

    def process_files(self):
        if not self.master_file_path or not self.target_folder:
            raise ValueError("请先选择 Master 文件和目标文件夹！")

        # 记录开始时间
        start_time = time.time()

        try:
            self.log("正在读取 Master 文件...")
            master_start_time = time.time()
            # 优化：只读取必要的列，并直接指定数据类型为字符串
            usecols = [1, self.match_column_index+1, self.content_column_index]  # 1是Key列(B列)
            master_df = pd.read_excel(
                self.master_file_path,
                engine='openpyxl',
                dtype={col: str for col in range(len(usecols))},  # 直接指定所有列为字符串类型
                keep_default_na=False,
                usecols=usecols
            )
            master_end_time = time.time()
            self.log(f"Master文件读取耗时: {master_end_time - master_start_time:.2f}秒")
        except Exception as e:
            raise Exception(f"读取 Master 文件失败：{e}")

        # 优化：直接在创建字典时处理数据，避免额外的循环
        master_dict = {}
        master_data = master_df.values
        for row in master_data:
            key = row[0].strip() if row[0] else ''  # 直接处理空值情况
            if key:  # 只处理非空key
                match_val = row[1] if row[1] else ''
                content_val = row[2] if row[2] else ''
                if match_val:  # 只存储有效的匹配值
                    master_dict[key] = [match_val, content_val]

        self.log(f"Master 中共找到 {len(master_dict)} 个有效 Key")
        
        # 添加调试日志，打印特定key的内容
        # debug_key1 = "4D03332141C5B492D7E97891939EDDFB"
        # debug_key2 = "SysPhotograph.WBP_Photograph_EdtPage.StrengthText,SysPhotograph"
        # if debug_key1 or debug_key2 in master_dict:
        #     self.log(f"Debug - Key '{debug_key1}' 的内容: {master_dict[debug_key1]}")
        #     self.log(f"Debug - Key '{debug_key2}' 的内容: {master_dict[debug_key2]}")
        # else:
        #     self.log(f"Debug - 未找到Key: {debug_key1}")

        # 收集目标文件
        
        file_paths = []
        for root, _, files in os.walk(self.target_folder):
            file_paths.extend(
                os.path.join(root, file)
                for file in files
                if file.lower().endswith(('.xlsx', '.xls'))
            )

        self.log(f"找到 {len(file_paths)} 个目标文件")

        # 优化：调整线程池大小以获得更好的性能
        process_start_time = time.time()
        max_workers = min(32, len(file_paths))  # 限制最大线程数
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(self._process_single_file, fp, master_dict) for fp in file_paths]
            updated_count = sum(future.result() for future in concurrent.futures.as_completed(futures))
        process_end_time = time.time()

        total_time = time.time() - start_time
        self.log(f"文件处理耗时: {process_end_time - process_start_time:.2f}秒")
        self.log(f"总耗时: {total_time:.2f}秒")
        self.log(f"处理完成，共更新 {updated_count} 处数据")

        # 添加后处理步骤
        self.log("开始后处理步骤...")
        self._post_process(file_paths)
        self.log("后处理步骤完成")

        return updated_count

    def _process_single_file(self, file_path, master_dict):
        updates = {}
        updated = 0
        
        try:
            # 使用openpyxl的只读模式读取文件
            wb = openpyxl.load_workbook(filename=file_path, read_only=True)
            ws = wb.active
            
            # 获取目标列的索引
            key_col = 'A'  # 第一列
            match_col = chr(ord('A') + self.match_column_index)  # 匹配列
            for idx, row in enumerate(ws.rows, start=1):
                try:
                    # 只读取需要的列
                    key_cell = row[0]
                    match_cell = row[self.match_column_index]
                    
                    # 确保单元格值转换为字符串
                    target_key = str(key_cell.value).strip() if key_cell.value else ''
                    target_match_value = str(match_cell.value) if match_cell.value else ''
                    
                    if not target_key or not target_match_value:
                        continue
                    
                    if target_key in master_dict:
                        master_values = master_dict[target_key]
                        if target_match_value == master_values[0]:
                            update_col = self.update_column_index + 1
                            updates[(idx, update_col)] = master_values[1]
                            updated += 1
                except Exception:
                    continue
            
            # 关闭只读工作簿
            wb.close()
            
            # 如果有更新，重新打开文件进行写入
            if updates:
                try:
                    wb = openpyxl.load_workbook(file_path)
                    ws = wb.active
                    
                    # 批量更新单元格
                    for (row, col), value in updates.items():
                        # 使用正确的方法获取和设置单元格值
                        cell = ws._get_cell(row, col)
                        if cell is None:
                            cell = ws._cell(row, col)
                        cell.value = value
                        
                    wb.save(file_path)
                except Exception:
                    return 0
                finally:
                    if wb:
                        wb.close()
                    
        except Exception as e:
            return 0
            
        return updated
        
    def _post_process(self, file_paths):
        """使用win32com.client处理Excel文件以确保兼容性，采用最简单的单线程处理方式"""
        try:
            post_process_start_time = time.time()
            total_files = len(file_paths)
            
            # 创建一个Excel实例供所有文件使用
            excel_app = Dispatch('Excel.Application')
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            try:
                # 简单循环处理每个文件
                for index, file_path in enumerate(file_paths, 1):
                    self.log(f"正在后处理文件 ({index}/{total_files}): {os.path.basename(file_path)}")
                    self._process_single_file_post(file_path, excel_app)

            finally:
                # 确保Excel实例被正确关闭和释放
                if excel_app is not None:
                    try:
                        excel_app.Quit()
                    except:
                        pass
                    excel_app = None
            
            post_process_end_time = time.time()
            self.log(f"后处理步骤耗时: {post_process_end_time - post_process_start_time:.2f}秒")
                    
        except Exception as e:
            self.log(f"后处理步骤失败：{str(e)}")
    
    def _process_single_file_post(self, file_path, excel_app):
        """处理单个文件的后处理逻辑，使用共享的Excel实例"""
        try:
            # 打开工作簿
            wb = excel_app.Workbooks.Open(file_path)
            if wb is not None:
                wb.Save()
                wb.Close(True)
                wb = None  # 显式释放工作簿对象
        except Exception as e:
            self.log(f"后处理文件 {os.path.basename(file_path)} 时出错：{str(e)}")
        finally:
            # 确保Excel实例被正确关闭和释放
            if excel_app is not None:
                try:
                    excel_app.Quit()
                    excel_app = None  # 显式释放Excel应用程序对象
                except:
                    pass