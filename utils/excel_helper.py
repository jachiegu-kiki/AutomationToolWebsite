import os
import pandas as pd
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def get_raw_files(raw_data_dir, header_keyword="条线"):
    """
    扫描目录并嗅探每个 Sheet 潜在的‘拆分依据’列及表头位置
    :param header_keyword: 用户指定的表头定位关键词 (默认为 "条线")
    """
    files_info = {}
    if not os.path.exists(raw_data_dir): return files_info

    for filename in os.listdir(raw_data_dir):
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            file_path = os.path.join(raw_data_dir, filename)
            try:
                # 使用 with 上下文管理器，确保读取后自动关闭档案释放资源
                with pd.ExcelFile(file_path) as xls:
                    sheets_data = {}
                    for sn in xls.sheet_names:
                        # 读取前20行进行表头嗅探
                        df_peek = pd.read_excel(xls, sheet_name=sn, header=None, nrows=20)
                        potential_cols = []
                        header_idx = 0

                        found_header = False
                        for i, row in df_peek.iterrows():
                            # 查找包含关键词的单元格作为潜在表头
                            # 使用 str() 转换防止非字符串类型报错
                            if any(header_keyword in str(val) for val in row.values if pd.notna(val)):
                                potential_cols = [str(c).strip() for c in row.values if pd.notna(c)]
                                header_idx = i
                                found_header = True
                                break

                        # 如果没找到关键词，默认第一行为表头（兜底策略）
                        if not found_header and not df_peek.empty:
                            potential_cols = [str(c).strip() for c in df_peek.iloc[0].values if pd.notna(c)]
                            header_idx = 0

                        sheets_data[sn] = {
                            "columns": potential_cols,
                            "header_idx": header_idx,
                            # 尝试自动选中包含关键词的列
                            "suggested_col": next((c for c in potential_cols if header_keyword in c), "")
                        }
                    files_info[filename] = sheets_data
            except Exception as e:
                logging.error(f"分析文件 {filename} 失败: {e}")
    return files_info


def process_and_split(selection, raw_data_dir, output_dir):
    """
    按用户指定的字段进行拆分
    selection 结构: { "文件名": { "原名": {"custom": "新名", "col": "拆分列", "h": 行号} } }
    """
    if not os.path.exists(output_dir): os.makedirs(output_dir)
    # 清理旧文件
    for f in os.listdir(output_dir):
        try:
            os.remove(os.path.join(output_dir, f))
        except:
            pass

    data_buffer = {}
    logs = []

    for filename, config in selection.items():
        file_path = os.path.join(raw_data_dir, filename)
        try:
            with pd.ExcelFile(file_path) as xls:
                for origin_sheet, info in config.items():
                    # 使用动态探测到的表头行号读取数据
                    df = pd.read_excel(xls, sheet_name=origin_sheet, header=info['h'])
                    df.columns = [str(c).strip() for c in df.columns]

                    target_col = info['col']
                    if target_col not in df.columns:
                        logs.append(f"错误: {filename}-{origin_sheet} 找不到列 [{target_col}]")
                        continue

                    # 移除拆分列为空的行
                    df = df.dropna(subset=[target_col])

                    for line_val, group_df in df.groupby(target_col):
                        line_name = str(line_val).strip()
                        if not line_name: continue # 跳过空名

                        if line_name not in data_buffer: data_buffer[line_name] = {}
                        # 处理同名 Sheet 覆盖问题：如果多个源 Sheet 拆分出同一个 line_name
                        # 这里简单处理为追加 Sheet 名后缀或直接使用 custom 名
                        sheet_out_name = info['custom']
                        # Excel Sheet 名最长31字符，需截断
                        sheet_out_name = sheet_out_name[:31]

                        data_buffer[line_name][sheet_out_name] = group_df

        except Exception as e:
            logs.append(f"处理 {filename} 异常: {str(e)}")

    for line_name, sheets in data_buffer.items():
        # 清洗文件名，防止非法字符
        safe_line = "".join([c for c in line_name if c not in r'\/*?:"<>|'])
        path = os.path.join(output_dir, f"{safe_line}.xlsx")
        try:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                for s_name, s_df in sheets.items():
                    s_df.to_excel(writer, sheet_name=s_name, index=False)
        except Exception as e:
            logs.append(f"写入文件 {safe_line}.xlsx 失败: {str(e)}")

    return {"status": "success", "message": f"拆分完毕，生成 {len(data_buffer)} 个文件", "logs": logs}