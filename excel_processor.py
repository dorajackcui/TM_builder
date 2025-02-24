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
            # 读取 Master 文件时就将所有列转换为字符串类型，只读取必要的列
            usecols = [1, self.match_column_index+1, self.update_column_index+1]  # 1是Key列(B列)
            master_df = pd.read_excel(
                self.master_file_path,
                engine='openpyxl',
                dtype=str,
                keep_default_na=False,
                usecols=usecols
            )
        except Exception as e:
            raise Exception(f"读取 Master 文件失败：{e}")

        # 标准化处理，确保所有列的数据类型一致性
        for col in master_df.columns:
            master_df[col] = master_df[col].apply(lambda x: str(x) if x is not None else 'None')

        # 创建数据结构：{Key: [匹配列值, 更新列值]}
        master_data = master_df.values
        master_dict = {}
        for row in master_data:
            key = row[0]  # Key列现在是第一列，因为我们只读取了需要的列
            match_val = row[1]  # 匹配列是第二列
            update_val = row[2]  # 更新列是第三列
            master_dict[key] = [match_val, update_val]
        self.log(f"Master 中共找到 {len(master_dict)} 个有效 Key")
        
        # 添加调试日志，打印特定key的内容
        debug_key1 = "4D03332141C5B492D7E97891939EDDFB"
        debug_key2 = "SysPhotograph.WBP_Photograph_EdtPage.StrengthText,SysPhotograph"
        if debug_key1 or debug_key2 in master_dict:
            self.log(f"Debug - Key '{debug_key1}' 的内容: {master_dict[debug_key1]}")
            self.log(f"Debug - Key '{debug_key2}' 的内容: {master_dict[debug_key2]}")
        else:
            self.log(f"Debug - 未找到Key: {debug_key1}")

        # 收集目标文件
        file_paths = []
        for root, _, files in os.walk(self.target_folder):
            for file in files:
                if file.lower().endswith(('.xlsx', '.xls')):
                    file_paths.append(os.path.join(root, file))
        self.log(f"找到 {len(file_paths)} 个目标文件")
        self.log(f"匹配列index： {self.match_column_index}")

        # 并行处理
        with concurrent.futures.ThreadPoolExecutor() as executor:
            results = executor.map(
                self._process_single_file,
                file_paths,
                [master_dict]*len(file_paths)
            )

        updated_count = sum(results)
        self.log(f"处理完成，共更新 {updated_count} 处数据")
        return updated_count

    def _process_single_file(self, file_path, master_dict):
        try:
            # 只读取需要的列（Key列、匹配列和更新列）
            usecols = [0, self.match_column_index, self.update_column_index]  # 0是Key列
            df = pd.read_excel(
                file_path,
                header=0,
                usecols=usecols,
                dtype=str,
                keep_default_na=False
            )
            
            # 标准化处理数据
            for col in df.columns:
                df[col] = df[col].apply(lambda x: str(x) if x is not None else 'None')
                
            # 加载工作簿以保持格式
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
        except Exception as e:
            return 0

        # 创建一个字典来存储需要更新的单元格位置和值
        updates = {}
        updated = 0
        
        for idx in df.index:
            try:
                # 获取并标准化目标文件数据
                target_key = str(df.iat[idx, 0]).strip()  # 第一列为Key
                target_match_value = str(df.iat[idx, 1])  # 匹配列（在读取时已经调整为第二列）
                
                # 空值检查
                if pd.isna(target_key) or target_key == "":
                    continue
                if target_match_value == "":
                    continue
                    
                # 匹配 Master 数据并更新
                if target_key in master_dict:
                    master_values = master_dict[target_key]
                    if target_match_value == master_values[0]:
                        # 获取实际的Excel列号
                        update_col = self.update_column_index + 1
                        updates[(idx + 2, update_col)] = master_values[1]
                        updated += 1
            except Exception as e:
                continue
                
        # 批量更新单元格
        for (row, col), value in updates.items():
            ws.cell(row=row, column=col).value = value
    
        try:
            # 保存工作簿
            wb.save(file_path)
        except Exception as e:
            return 0
        return updated