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
            master_df = pd.read_excel(self.master_file_path)
        except Exception as e:
            raise Exception(f"读取 Master 文件失败：{e}")

        # 预处理 Master 文件
        master_df = master_df.drop(master_df.columns[0], axis=1)  # 删除 A 列
        if "Key" not in master_df.columns:
            raise ValueError("Master 文件中没有找到 'Key' 列！")

        # 标准化处理，确保所有列的数据类型一致性
        for col in master_df.columns:
            # 将所有列转换为字符串并清理空白字符
            if col != 'Key':  # 对非Key列保持原始数据类型
                master_df[col] = master_df[col].apply(lambda x: '' if pd.isna(x) else str(x).strip())
            else:
                master_df[col] = master_df[col].astype(str).str.strip()

        # 创建数据结构：{Key: [中文值, 列2值, 列3值...]}
        self.master_columns = master_df.columns.tolist()[1:]  # 存储列顺序（排除Key列）
        master_data = master_df.values
        master_dict = {
            row[0]: list(row[1:])  # [中文值, 列2值, 列3值...]
            for row in master_data
        }
        self.log(f"Master 中共找到 {len(master_dict)} 个有效 Key")

        # 收集目标文件
        file_paths = []
        for root, _, files in os.walk(self.target_folder):
            for file in files:
                if file.lower().endswith(('.xlsx', '.xls')):
                    file_paths.append(os.path.join(root, file))
        self.log(f"找到 {len(file_paths)} 个目标文件")

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
            # 使用 openpyxl 引擎读取文件以保留格式
            df = pd.read_excel(file_path, header=0, engine='openpyxl')
            # 同时打开原始文件以保留格式
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
        except Exception as e:
            self.log(f"文件读取失败：{file_path}\n错误：{str(e)}")
            return 0
    
        updated = 0
        for idx in df.index:
            # 获取并标准化目标文件数据
            try:
                target_key = str(df.iat[idx, 0]).strip()  # 第一列为Key
                target_chinese = str(df.iat[idx, 1]).strip()  # 第二列为中文列
                
                # 空值检查
                if pd.isna(target_key) or target_key.lower() == "nan" or target_key == "":
                    continue
                if pd.isna(target_chinese) or target_chinese.lower() == "nan" or target_chinese == "":
                    continue
            except Exception as e:
                self.log(f"数据读取错误：第 {idx+1} 行 - {str(e)}")
                continue
    
            # 匹配 Master 数据
            if target_key in master_dict:
                master_values = master_dict[target_key]
                
                # 检查中文列是否匹配
                if target_chinese == master_values[0]:
                    # 按列位置复制数据，保持数据类型
                    for col_offset, value in enumerate(master_values[1:], start=2):
                        if col_offset >= df.shape[1]:
                            break
                        try:
                            # 尝试保持原始数据类型
                            original_value = df.iat[idx, col_offset]
                            if pd.isna(original_value):
                                if pd.isna(value):
                                    df.iat[idx, col_offset] = pd.NA
                                else:
                                    # 根据原始列的数据类型进行转换
                                    try:
                                        if df[df.columns[col_offset]].dtype in ['int64', 'float64']:
                                            df.iat[idx, col_offset] = float(value) if '.' in str(value) else int(value)
                                        else:
                                            df.iat[idx, col_offset] = str(value)
                                    except (ValueError, TypeError):
                                        df.iat[idx, col_offset] = str(value)
                            else:
                                # 如果原始值是数字类型，尝试转换新值为相同类型
                                if isinstance(original_value, (int, float)):
                                    try:
                                        df.iat[idx, col_offset] = float(value) if '.' in str(value) else int(value)
                                    except (ValueError, TypeError):
                                        df.iat[idx, col_offset] = str(value)
                                else:
                                    df.iat[idx, col_offset] = str(value) if value is not None else ''
                        except Exception as e:
                            self.log(f"数据转换错误：第 {idx+1} 行，第 {col_offset+1} 列 - {str(e)}")
                            continue
                    updated += 1
                    # self.log(f"文件 {os.path.basename(file_path)} 第 {idx+1} 行已更新")
                else:
                    # self.log(f"文件 {os.path.basename(file_path)} 第 {idx+1} 行中文列不匹配")
                    # self.log(f"文件 {os.path.basename(file_path)} 第 {idx+1} 行 Key 不存在")
                    pass
    
        try:
            # 将更新后的数据写回到原始工作表中，保留格式
            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    ws.cell(row=row_idx + 2, column=col_idx + 1).value = df.iat[row_idx, col_idx]
            
            # 保存工作簿
            wb.save(file_path)
        except Exception as e:
            self.log(f"文件保存失败：{file_path}\n错误：{str(e)}")
            return 0
        return updated