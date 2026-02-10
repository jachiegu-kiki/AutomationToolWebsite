import os
import pandas as pd
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def get_raw_files(raw_data_dir):
    """扫描目录并嗅探每个 Sheet 潜在的'条线'列及表头位置"""
    files_info = {}
    if not os.path.exists(raw_data_dir):
        return files_info

    for filename in os.listdir(raw_data_dir):
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            file_path = os.path.join(raw_data_dir, filename)
            try:
                with pd.ExcelFile(file_path) as xls:
                    sheets_data = {}
                    for sn in xls.sheet_names:
                        df_peek = pd.read_excel(xls, sheet_name=sn, header=None, nrows=20)
                        potential_cols = []
                        header_idx = 0

                        for i, row in df_peek.iterrows():
                            if any("条线" in str(val) for val in row.values if pd.notna(val)):
                                potential_cols = [str(c).strip() for c in row.values if pd.notna(c)]
                                header_idx = i
                                break

                        sheets_data[sn] = {
                            "columns": potential_cols,
                            "header_idx": header_idx,
                            "suggested_col": next((c for c in potential_cols if "条线" in c), "")
                        }
                    files_info[filename] = sheets_data
            except Exception as e:
                logging.error(f"分析文件 {filename} 失败: {e}")
    return files_info


def clear_directory(dir_path):
    """清空指定目录下的所有文件（保留目录本身）"""
    removed = 0
    failed = 0
    if not os.path.exists(dir_path):
        return {"removed": 0, "failed": 0}

    for filename in os.listdir(dir_path):
        file_path = os.path.join(dir_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                removed += 1
        except Exception as e:
            logging.error(f"删除文件 {filename} 失败: {e}")
            failed += 1

    return {"removed": removed, "failed": failed}


# ==============================================================
#  分组映射模块 (整合自 split_by_tiaoxian.py)
#
#  映射分为两张独立的表：
#    「签约」← 附件1签约明细表   (本期 + 同期 合并)
#    「退费」← 附件4退费及转国家明细 (本期 + 同期 合并)
# ==============================================================

# Sheet 名 → 映射类别
_SHEET_TO_CATEGORY = {
    "附件1签约明细表":     "签约",
    "附件4退费及转国家明细": "退费",
}


def _safe_read_excel(file_path, sheet_name=None, header=None, **kwargs):
    """安全读取 Excel，自动尝试多种引擎"""
    engines = ['calamine', 'openpyxl', 'xlrd']
    last_error = None
    for engine in engines:
        try:
            return pd.read_excel(
                file_path, sheet_name=sheet_name,
                header=header, engine=engine, **kwargs
            )
        except Exception as e:
            last_error = e
            continue
    raise Exception(f"无法读取文件 {file_path}，所有引擎均失败: {last_error}")


def _safe_excel_file(file_path):
    """安全创建 ExcelFile 对象"""
    engines = ['calamine', 'openpyxl']
    last_error = None
    for engine in engines:
        try:
            return pd.ExcelFile(file_path, engine=engine)
        except Exception as e:
            last_error = e
            continue
    raise Exception(f"无法创建 ExcelFile: {file_path}, {last_error}")


def build_group_mapping(ref_config):
    """
    从参照文件构建两张独立的「合同号 → 分组部门」映射字典。

    参照文件结构 (本期 / 同期各一份):
      Sheet「附件1签约明细表」     → 签约映射
      Sheet「附件4退费及转国家明细」 → 退费映射

    本期与同期的同类 Sheet 数据合并到同一张映射表中。

    :param ref_config: {
        "ref_benqi":  "本期参照路径",
        "ref_tongqi": "同期参照路径"
    }
    :return: {
        "mappings": { "签约": {...}, "退费": {...} },
        "logs": [...],
        "stats": { "签约_benqi", "签约_tongqi", "签约_total",
                   "退费_benqi", "退费_tongqi", "退费_total", "total" }
    }
    """
    mappings = {"签约": {}, "退费": {}}
    logs = []
    stats = {
        "签约_benqi": 0, "签约_tongqi": 0, "签约_total": 0,
        "退费_benqi": 0, "退费_tongqi": 0, "退费_total": 0,
        "total": 0,
    }

    required_cols = ["合同号", "分组部门"]

    for ref_key, ref_label in [("ref_benqi", "本期"), ("ref_tongqi", "同期")]:
        ref_path = ref_config.get(ref_key, "").strip()
        if not ref_path:
            logs.append(f"[跳过] {ref_label} 路径未配置")
            continue

        if not os.path.exists(ref_path):
            logs.append(f"[错误] {ref_label} 文件不存在: {ref_path}")
            continue

        try:
            xls = _safe_excel_file(ref_path)
            available_sheets = xls.sheet_names

            for sheet_name, category in _SHEET_TO_CATEGORY.items():
                if sheet_name not in available_sheets:
                    logs.append(f"[跳过] {ref_label} 中无 Sheet: {sheet_name}")
                    continue

                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
                    df.columns = [str(c).strip() for c in df.columns]

                    missing = [c for c in required_cols if c not in df.columns]
                    if missing:
                        logs.append(
                            f"[警告] {ref_label}-{sheet_name} 缺少列 {missing}，"
                            f"可用列: {list(df.columns)[:10]}"
                        )
                        continue

                    valid = df.dropna(subset=required_cols)
                    count_before = len(mappings[category])

                    for contract_id, group_dept in zip(valid["合同号"], valid["分组部门"]):
                        mappings[category][contract_id] = str(group_dept).strip()

                    added = len(mappings[category]) - count_before
                    suffix = ref_key.replace("ref_", "")  # "benqi" / "tongqi"
                    stats[f"{category}_{suffix}"] = added

                    logs.append(
                        f"[成功] {ref_label}-{sheet_name} → {category}映射: "
                        f"读取 {len(valid)} 行, 新增 {added} 条"
                    )
                except Exception as e:
                    logs.append(f"[错误] 读取 {ref_label}-{sheet_name} 失败: {e}")

            try:
                xls.close()
            except:
                pass

        except Exception as e:
            logs.append(f"[错误] 打开 {ref_label} 参照文件失败: {e}")

    stats["签约_total"] = len(mappings["签约"])
    stats["退费_total"] = len(mappings["退费"])
    stats["total"] = stats["签约_total"] + stats["退费_total"]

    logs.append(
        f"[完成] 签约映射: {stats['签约_total']} 条, "
        f"退费映射: {stats['退费_total']} 条, "
        f"合计: {stats['total']} 条"
    )

    return {"mappings": mappings, "logs": logs, "stats": stats}


# ==============================================================
#  拆分核心 (支持按 Sheet 选择 签约/退费 分组)
# ==============================================================

def process_and_split(selection, raw_data_dir, output_dir, group_mappings=None):
    """
    按用户指定的字段进行拆分，并根据每个 Sheet 的分组类别追加「分组」字段。

    :param selection: {
        "文件名": {
            "Sheet原名": {
                "custom":     "输出Sheet名",
                "col":        "拆分列",
                "h":          表头行号,
                "group_type": "签约" | "退费" | ""
            }
        }
    }
    :param group_mappings: { "签约": {合同号: 分组}, "退费": {合同号: 分组} }
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 清理旧文件
    for f in os.listdir(output_dir):
        try:
            os.remove(os.path.join(output_dir, f))
        except:
            pass

    data_buffer = {}
    logs = []

    if group_mappings:
        qy = len(group_mappings.get("签约", {}))
        tf = len(group_mappings.get("退费", {}))
        logs.append(f"✅ 已加载分组映射 — 签约: {qy} 条, 退费: {tf} 条")
    else:
        logs.append("ℹ️ 未提供分组映射，跳过「分组」字段追加")

    for filename, config in selection.items():
        file_path = os.path.join(raw_data_dir, filename)
        try:
            with pd.ExcelFile(file_path) as xls:
                for origin_sheet, info in config.items():
                    df = pd.read_excel(xls, sheet_name=origin_sheet, header=info['h'])
                    df.columns = [str(c).strip() for c in df.columns]

                    # ====== 按 group_type 选择对应映射表追加「分组」======
                    group_type = info.get('group_type', '')
                    if group_mappings and group_type and group_type in group_mappings:
                        chosen_map = group_mappings[group_type]
                        if "合同号" in df.columns:
                            df["分组"] = df["合同号"].map(chosen_map)
                            matched = df["分组"].notna().sum()
                            total = len(df)
                            logs.append(
                                f"  [{filename}-{origin_sheet}] "
                                f"使用「{group_type}」映射, 命中 {matched}/{total} 行"
                            )
                        else:
                            logs.append(
                                f"  [{filename}-{origin_sheet}] "
                                f"无「合同号」列, 无法套用「{group_type}」映射"
                            )
                    elif group_type:
                        logs.append(
                            f"  [{filename}-{origin_sheet}] "
                            f"指定了「{group_type}」映射但映射表未加载, 跳过"
                        )
                    # ====== 追加逻辑结束 ======

                    target_col = info['col']
                    if target_col not in df.columns:
                        logs.append(f"错误: {filename}-{origin_sheet} 找不到列 [{target_col}]")
                        continue

                    df = df.dropna(subset=[target_col])
                    for line_val, group_df in df.groupby(target_col):
                        line_name = str(line_val).strip()
                        if line_name not in data_buffer:
                            data_buffer[line_name] = {}
                        data_buffer[line_name][info['custom']] = group_df

        except Exception as e:
            logs.append(f"处理 {filename} 异常: {str(e)}")

    for line_name, sheets in data_buffer.items():
        safe_line = "".join([c for c in line_name if c not in r'\/*?:"<>|'])
        path = os.path.join(output_dir, f"{safe_line}.xlsx")
        try:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                for s_name, s_df in sheets.items():
                    s_df.to_excel(writer, sheet_name=s_name, index=False)
        except Exception as e:
            logs.append(f"写入文件 {safe_line}.xlsx 失败: {str(e)}")

    return {
        "status": "success",
        "message": f"拆分完毕，生成 {len(data_buffer)} 个文件",
        "logs": logs
    }
