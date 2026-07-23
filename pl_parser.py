"""
Packing List 解析模块
负责从上传的 PL Excel 文件中提取 SKU / 品名 (Description) / HTS Code / 税率 四元组。

关键处理点：
1. PL 格式不固定（表头行位置、列名可能变化），需要自动定位表头行。
2. 同一产品多个 SKU 时，品名/HTS 经常只写在第一行，下面的行是空白（"继承"上一行的值），
   必须做 forward-fill，否则会被误判为"品名/HTS 缺失"或产生噪音预警。
3. 自动跳过 TOTAL / 合计行、空行、签名行等非数据行。
4. 税率列：很多 PL 模板里这一列没有文字表头（纯数字列），所以不能靠关键词识别。
   采用"位置优先 + 自动侦测兜底"的策略，找不到就整列留空，不报错、不预警。
"""

import pandas as pd
import re


# 常见列名关键词（不区分大小写），用于自动识别表头
COLUMN_KEYWORDS = {
    "sku": ["sku"],
    "description": ["item description", "description", "product name", "品名", "货描", "item name"],
    "hts": ["hts", "hs code", "hscode", "h.t.s", "h.s.code", "海关编码", "商品编码"],
    "tax_rate": ["tax rate", "duty rate", "税率", "关税税率", "duty%", "tax%"],
}

# 税率列没有文字表头时，按经验位置在 SKU 列右边第几列去尝试（0-indexed 偏移）
TAX_RATE_POSITION_OFFSET = 2

# 自动侦测税率列时，一列里"看起来像税率"（0~100 的数字）的行占比至少要达到这个比例
TAX_RATE_DETECT_MIN_RATIO = 0.6


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
    根据表头行的实际文本，把每一列映射到标准字段名 (sku / description / hts / tax_rate)。
    注意：税率列在多数模板里没有文字表头，所以这里通常映射不到 tax_rate，
    需要靠 _locate_tax_rate_column 按位置/自动侦测补充。
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


def _parse_tax_value(val) -> float | None:
    """把单个单元格解析成 0~100 的税率数字，解析不了或超出范围返回 None。"""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        f = float(val)
    else:
        text = str(val).strip().replace("%", "").replace(",", ".")
        if text == "":
            return None
        try:
            f = float(text)
        except ValueError:
            return None
    if 0 <= f <= 100:
        return f
    return None


def _column_looks_like_tax_rate(series: pd.Series) -> bool:
    """判断一列是否"看起来像税率"：多数非空值是 0~100 的数字。"""
    non_null = series.dropna()
    non_null = non_null[non_null.astype(str).str.strip() != ""]
    if len(non_null) == 0:
        return False
    valid = sum(1 for v in non_null if _parse_tax_value(v) is not None)
    return (valid / len(non_null)) >= TAX_RATE_DETECT_MIN_RATIO


def _locate_tax_rate_column(raw: pd.DataFrame, header_idx: int, sku_col_idx: int,
                             already_mapped_cols: set) -> int | None:
    """
    定位税率列，策略：位置优先，自动侦测兜底。
    1. 先尝试 SKU 列右边第 TAX_RATE_POSITION_OFFSET 列（常见模板的经验位置）。
    2. 不满足就往右扫描其余未被占用的列，找第一个"看起来像税率"的列。
    3. 都找不到就返回 None（调用方留空，不报警）。
    """
    data_rows = raw.iloc[header_idx + 1:]
    total_cols = raw.shape[1]

    candidate = sku_col_idx + TAX_RATE_POSITION_OFFSET
    if candidate < total_cols and candidate not in already_mapped_cols:
        if _column_looks_like_tax_rate(data_rows[candidate]):
            return candidate

    for col_idx in range(sku_col_idx + 1, total_cols):
        if col_idx in already_mapped_cols or col_idx == candidate:
            continue
        if _column_looks_like_tax_rate(data_rows[col_idx]):
            return col_idx

    return None


def _format_tax_rate(val) -> str:
    """把税率数字格式化成百分比字符串，如 6.5 -> '6.5%'，0 -> '0%'。找不到值返回空字符串。"""
    f = _parse_tax_value(val)
    if f is None:
        return ""
    if f == int(f):
        return f"{int(f)}%"
    return f"{f}%"


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

    # ---- 2b. 定位税率列（大多数模板没有文字表头，位置优先+自动侦测兜底） ----
    tax_col_idx = None
    if "tax_rate" in found_fields:
        # 表头文字里就写了"税率"/"tax rate"等关键词，直接用
        tax_col_idx = [c for c, f in col_map.items() if f == "tax_rate"][0]
    else:
        sku_col_idx = [c for c, f in col_map.items() if f == "sku"][0]
        already_mapped = set(col_map.keys())
        tax_col_idx = _locate_tax_rate_column(raw, header_idx, sku_col_idx, already_mapped)

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
    if tax_col_idx is not None:
        extracted["tax_rate_raw"] = data[tax_col_idx]
    else:
        extracted["tax_rate_raw"] = pd.NA
        warnings.append("未在文件中找到税率列，本次解析的税率将留空。")

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

    # 税率同样按"同一产品多 SKU 继承上一行"的逻辑 forward-fill（找到税率列时才处理）
    if tax_col_idx is not None:
        extracted["tax_rate_raw"] = extracted["tax_rate_raw"].replace(r"^\s*$", pd.NA, regex=True)
        extracted["tax_rate_raw"] = extracted["tax_rate_raw"].ffill()

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

    # 税率格式化为百分比字符串，如 6.5 -> "6.5%"；没有税率列或解析不了的留空，不报警
    extracted["tax_rate"] = extracted["tax_rate_raw"].apply(_format_tax_rate)
    extracted = extracted.drop(columns=["tax_rate_raw"])

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
        "records": extracted[["sku", "description", "hts", "tax_rate"]],
        "warnings": warnings,
    }
