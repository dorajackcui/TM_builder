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
            master_df = pd.read_excel(self.master_file_path, keep_default_na=False)
        except Exception as e:
            raise Exception(f"读取 Master 文件失败：{e}")

        # 预处理 Master 文件
        master_df = master_df.drop(master_df.columns[0], axis=1)  # 删除 A 列
        if "Key" not in master_df.columns:
            raise ValueError("Master 文件中没有找到 'Key' 列！")

        # 标准化处理，确保所有列的数据类型一致性
        for col in master_df.columns:
            # 将所有列转换为字符串，保留None字符串
            master_df[col] = master_df[col].apply(lambda x: str(x) if x is not None else 'None')

        # 创建数据结构：{Key: [中文值, 列2值, 列3值...]}
        self.master_columns = master_df.columns.tolist()[1:]  # 存储列顺序（排除Key列）
        master_data = master_df.values
        master_dict = {
            row[0]: list(row[1:])  # [中文值, 列2值, 列3值...]
            for row in master_data
        }
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
            # 读取文件时就将所有列转换为字符串类型
            df = pd.read_excel(file_path, header=0, engine='openpyxl', dtype=str, keep_default_na=False)
            
            # 添加与 master 文件相同的数据处理逻辑
            for col in df.columns:
                df[col] = df[col].apply(lambda x: str(x) if x is not None else 'None')
                
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
        except Exception as e:
            return 0

        updated = 0
        for idx in df.index:
            # 获取并标准化目标文件数据
            try:
                # 检查列索引是否超出范围
                if self.match_column_index >= df.shape[1] or self.update_column_index >= df.shape[1]:
                    continue
                
                # 检查第一列（Key列）是否存在
                if df.shape[1] == 0:
                    continue
                    
                target_key = str(df.iat[idx, 0]).strip()  # 第一列为Key
                target_match_value = str(df.iat[idx, self.match_column_index])  # 用户指定的匹配列
                
                # 空值检查
                if pd.isna(target_key) or target_key == "":
                    continue
                # 修改匹配列的空值检查逻辑，只有真正的空值和空字符串才跳过
                if target_match_value == "":
                    continue
            except Exception as e:
                continue
    
            # 匹配 Master 数据
            if target_key in master_dict:
                master_values = master_dict[target_key]
                
                # 检查匹配列是否匹配
                if target_match_value == master_values[self.match_column_index - 1]:
                    try:
                        # 只更新指定的列
                        update_value = master_values[self.update_column_index - 1]
                        df.iat[idx, self.update_column_index] = update_value
                        updated += 1
                    except Exception as e:
                        continue
    
        try:
            # 将更新后的数据写回到原始工作表中，保留格式
            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    cell_value = df.iat[row_idx, col_idx]
                    cell = ws.cell(row=row_idx + 2, column=col_idx + 1)
                    cell.value = cell_value
            
            # 保存工作簿
            wb.save(file_path)
        except Exception as e:
            return 0
        return updated