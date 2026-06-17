"""
Packing List 解析模块
负责从上传的 PL Excel 文件中提取 SKU / 品名 (Description) / HTS Code 三元组。

关键处理点：
1. PL 格式不固定（表头行位置、列名可能变化），需要自动定位表头行。
2. 同一产品多个 SKU 时，品名/HTS 经常只写在第一行，下面的行是空白（"继承"上一行的值），
   必须做 forward-fill，否则会被误判为"品名/HTS 缺失"或产生噪音预警。
3. 自动跳过 TOTAL / 合计行、空行、签名行等非数据行。
"""

import pandas as pd
import re


# 常见列名关键词（不区分大小写），用于自动识别表头
COLUMN_KEYWORDS = {
    "sku": ["sku"],
    "description": ["item description", "description", "product name", "品名", "货描", "item name"],
    "hts": ["hts", "hs code", "hscode", "h.t.s", "h.s.code", "海关编码", "商品编码"],
}


def _find_header_row(raw_df: pd.DataFrame) -> int | None:
    """
    扫描原始 DataFrame（无表头读入），找到同时包含 SKU / Description / HTS
    关键词的那一行，作为真正的表头行索引。
    """
    max_scan_rows = min(50, len(raw_df))
    for row_idx in range(max_scan_rows):
        row_values = [str(v).lower() for v in raw_df.iloc[row_idx].tolist()]
        row_text = " | ".join(row_values)

        has_sku = any(kw in row_text for kw in COLUMN_KEYWORDS["sku"])
        has_desc = any(kw in row_text for kw in COLUMN_KEYWORDS["description"])
        has_hts = any(kw in row_text for kw in COLUMN_KEYWORDS["hts"])

        if has_sku and has_desc and has_hts:
            return row_idx
    return None


def _map_columns(header_row: pd.Series) -> dict:
    """
    根据表头行的实际文本，把每一列映射到标准字段名 (sku / description / hts)。
    """
    mapping = {}
    for col_idx, raw_name in header_row.items():
        name_lower = str(raw_name).strip().lower()
        for field, keywords in COLUMN_KEYWORDS.items():
            if any(kw in name_lower for kw in keywords):
                # 避免重复映射（例如已经映射过 sku 就不要再覆盖）
                if field not in mapping.values():
                    mapping[col_idx] = field
                break
    return mapping


def parse_packing_list(file_obj_or_path, po_number_hint: str | None = None) -> dict:
    """
    解析 packing list 文件，返回标准化结果。

    Returns:
        {
            "success": bool,
            "error": str | None,
            "po_number": str | None,         # 自动提取的 PO/箱单号，找不到则为 None
            "records": pd.DataFrame,         # columns: sku, description, hts
            "warnings": list[str],           # 解析过程中的非阻断性提示
        }
    """
    warnings = []

    try:
        raw = pd.read_excel(file_obj_or_path, header=None, sheet_name=0)
    except Exception as e:
        return {"success": False, "error": f"无法读取 Excel 文件：{e}", "po_number": None,
                "records": pd.DataFrame(), "warnings": warnings}

    # ---- 1. 尝试提取 PO / 箱单号（用于来源追溯） ----
    po_number = po_number_hint
    if po_number is None:
        for row_idx in range(min(20, len(raw))):
            row_text = " ".join(str(v) for v in raw.iloc[row_idx].tolist())
            match = re.search(r"(PO\s*NO\.?|箱单号|排柜单号)\s*[:：]?\s*([A-Za-z0-9\-_/]+)", row_text, re.IGNORECASE)
            if match:
                po_number = match.group(2).strip()
                break

    # ---- 2. 定位表头行 ----
    header_idx = _find_header_row(raw)
    if header_idx is None:
        return {"success": False, "error": "未能在文件中找到包含 SKU / Description / HTS 的表头行，请检查文件格式。",
                "po_number": po_number, "records": pd.DataFrame(), "warnings": warnings}

    header_row = raw.iloc[header_idx]
    col_map = _map_columns(header_row)

    required_fields = {"sku", "description", "hts"}
    found_fields = set(col_map.values())
    missing = required_fields - found_fields
    if missing:
        return {"success": False, "error": f"表头中缺少必要列：{', '.join(missing)}",
                "po_number": po_number, "records": pd.DataFrame(), "warnings": warnings}

    # ---- 3. 提取数据区（表头下一行开始，到 TOTAL 行或空 SKU 截止） ----
    data = raw.iloc[header_idx + 1:].copy()
    data = data.reset_index(drop=True)

    # 反向映射：标准字段名 -> 原始列索引
    field_to_col = {v: k for k, v in col_map.items()}
    extracted = pd.DataFrame({
        "sku": data[field_to_col["sku"]],
        "description": data[field_to_col["description"]],
        "hts": data[field_to_col["hts"]],
    })

    # 去掉 TOTAL / 合计 等汇总行，以及完全空白行
    def is_summary_or_blank(row) -> bool:
        sku_val = str(row["sku"]).strip().lower()
        if sku_val in ("nan", "none", ""):
            # 没有 SKU 但有品名也可能是合计行说明文字，一并跳过
            return True
        if "total" in sku_val or "合计" in sku_val or "总计" in sku_val:
            return True
        return False

    extracted = extracted[~extracted.apply(is_summary_or_blank, axis=1)].reset_index(drop=True)

    if extracted.empty:
        return {"success": False, "error": "表头下方未找到任何有效数据行。",
                "po_number": po_number, "records": pd.DataFrame(), "warnings": warnings}

    # ---- 4. 关键：forward-fill description 和 hts ----
    # PL 中同一产品的后续 SKU 行通常不重复填写品名/HTS，需要向上继承。
    before_fill_desc_blanks = extracted["description"].isna().sum()
    before_fill_hts_blanks = extracted["hts"].isna().sum()

    extracted["description"] = extracted["description"].replace(r"^\s*$", pd.NA, regex=True)
    extracted["hts"] = extracted["hts"].replace(r"^\s*$", pd.NA, regex=True)
    extracted["description"] = extracted["description"].ffill()
    extracted["hts"] = extracted["hts"].ffill()

    if before_fill_desc_blanks > 0 or before_fill_hts_blanks > 0:
        warnings.append(
            f"检测到 {before_fill_desc_blanks} 行品名为空、{before_fill_hts_blanks} 行 HTS 为空，"
            f"已自动向上继承上一行的值（同一产品多 SKU 的常见格式）。"
        )

    # ---- 5. 规范化数据类型 ----
    extracted["sku"] = extracted["sku"].astype(str).str.strip()

    def normalize_hts(val) -> str:
        if pd.isna(val):
            return ""
        if isinstance(val, float) and val == int(val):
            return str(int(val))
        return str(val).strip()

    extracted["hts"] = extracted["hts"].apply(normalize_hts)
    extracted["description"] = extracted["description"].astype(str).str.strip()

    # 仍有缺失 HTS/description 的行，单独提示（forward-fill 后仍为空，说明文件第一行就缺失）
    still_missing = extracted[(extracted["hts"] == "") | (extracted["description"] == "") | (extracted["description"].str.lower() == "nan")]
    if not still_missing.empty:
        skus = ", ".join(still_missing["sku"].tolist())
        warnings.append(f"以下 SKU 在 forward-fill 后仍缺少品名或 HTS，请人工检查：{skus}")

    # 去除同一份 PL 内部重复的 SKU（同一票货里 SKU 不应重复，但保险起见去重）
    duplicated_in_file = extracted[extracted.duplicated(subset=["sku"], keep=False)]
    if not duplicated_in_file.empty:
        dup_skus = duplicated_in_file["sku"].unique().tolist()
        warnings.append(f"本文件内部发现重复 SKU（已保留首次出现）：{', '.join(dup_skus)}")
        extracted = extracted.drop_duplicates(subset=["sku"], keep="first").reset_index(drop=True)

    return {
        "success": True,
        "error": None,
        "po_number": po_number,
        "records": extracted[["sku", "description", "hts"]],
        "warnings": warnings,
    }
