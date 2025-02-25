import pandas as pd
import os
import concurrent.futures
import openpyxl

class ExcelProcessor:
    def __init__(self, log_callback=None):
        self.master_file_path = ""
        self.target_folder = ""
        self.log_callback = log_callback or (lambda msg: None)
        self.master_columns = []  # 存储列位置信息
        self.match_column_index = 1  # 默认使用第二列作为匹配列
        self.update_column_index = 2  # 默认更新第三列

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
            # 优化：只读取必要的列，并直接指定数据类型为字符串
            usecols = [1, self.match_column_index+1, self.update_column_index+1]  # 1是Key列(B列)
            master_df = pd.read_excel(
                self.master_file_path,
                engine='openpyxl',
                dtype={col: str for col in range(len(usecols))},  # 直接指定所有列为字符串类型
                keep_default_na=False,
                usecols=usecols
            )
        except Exception as e:
            raise Exception(f"读取 Master 文件失败：{e}")

        # 优化：直接在创建字典时处理数据，避免额外的循环
        master_dict = {}
        master_data = master_df.values
        for row in master_data:
            key = row[0].strip() if row[0] else ''  # 直接处理空值情况
            if key:  # 只处理非空key
                match_val = row[1] if row[1] else ''
                update_val = row[2] if row[2] else ''
                if match_val:  # 只存储有效的匹配值
                    master_dict[key] = [match_val, update_val]

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
        max_workers = min(32, len(file_paths))  # 限制最大线程数
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(self._process_single_file, fp, master_dict) for fp in file_paths]
            updated_count = sum(future.result() for future in concurrent.futures.as_completed(futures))

        self.log(f"处理完成，共更新 {updated_count} 处数据")
        return updated_count

    def _process_single_file(self, file_path, master_dict):
        try:
            # 优化：只读取必要的列，并直接指定数据类型
            usecols = [0, self.match_column_index, self.update_column_index]
            df = pd.read_excel(
                file_path,
                header=0,
                usecols=usecols,
                dtype={col: str for col in range(len(usecols))},
                keep_default_na=False
            )
            
            # 优化：使用向量化操作处理数据
            df = df.fillna('')  # 替换所有NaN为空字符串
            df = df.astype(str)  # 确保所有数据为字符串类型
            
            # 加载工作簿以保持格式
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
        except Exception as e:
            return 0

        updates = {}
        updated = 0
        
        # 优化：批量处理数据
        df_array = df.values
        for idx, row in enumerate(df_array):
            try:
                target_key = row[0].strip()
                target_match_value = row[1]
                
                if not target_key or not target_match_value:
                    continue
                
                if target_key in master_dict:
                    master_values = master_dict[target_key]
                    if target_match_value == master_values[0]:
                        update_col = self.update_column_index + 1
                        updates[(idx + 2, update_col)] = master_values[1]
                        updated += 1
            except Exception:
                continue

        # 批量更新单元格
        if updates:
            for (row, col), value in updates.items():
                ws.cell(row=row, column=col).value = value
            try:
                wb.save(file_path)
            except Exception:
                return 0

        return updated