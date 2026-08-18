"""
Packing List 解析模块
负责从上传的 PL Excel 文件中提取 SKU / 品名 (Description) / HTS Code / 税率 四元组。

关键处理点：
1. PL 格式不固定（表头行位置、列名可能变化），需要自动定位表头行。
2. 同一产品多个 SKU 时，品名/HTS 经常只写在第一行，下面的行是空白（"继承"上一行的值），
   必须做 forward-fill，否则会被误判为"品名/HTS 缺失"或产生噪音预警。
3. 自动跳过 TOTAL / 合计行、空行、签名行等非数据行。
4. 税率列：很多 PL 模板里这一列没有文字表头（纯数字列），所以不能靠关键词识别，
   而且往往紧挨着一列"税额/税金"（金额，而不是税率），两列数值都落在能通过
   "像百分比"检测的范围里。采用"候选列中取值最容易重复的一列"的策略
   （见 _locate_tax_rate_column）：税率本质上是按 HTS 分类挂钩的，取值范围
   很有限，会在同一批货里大量重复；税额是数量×单价×税率算出来的连续型金额，
   几乎每行都不同。找不到候选列时整列留空，不报错、不预警。
5. 有些文件是「发票 / 箱单 / 合同」三个表格放在同一个 xlsx 的不同 sheet 里
   （新增支持的格式），这种情况下必须只解析「箱单」(Packing List) 这个 sheet——
   发票、合同两个 sheet 的行结构完全不同，读它们会得到错误数据甚至解析失败。
   单 sheet 的旧格式文件不受影响，会自动退回读第一个 sheet。
6. SKU 列有时候表头是空白的（模板里就没写"SKU"这几个字），这种情况下改用数据
   特征来定位：SKU 几乎每一行都有值（哪怕品名/HTS 因为同产品多 SKU 只在第一行
   填写），而且基本不是纯数字——在 HTS 列右边找符合这两个特征的第一列。
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

# 自动侦测税率列时，一列里"看起来像税率"（0~100 的数字）的行占比至少要达到这个比例
TAX_RATE_DETECT_MIN_RATIO = 0.6

# 自动侦测税率列时，候选列"不重复值占比"（_distinct_value_ratio）不能超过这个上限。
# 税率本质上是按 HTS 分类挂钩的，取值范围很有限，同一批货里会大量重复；
# 而"每个 SKU 的件数/箱数"这类数量型列，虽然数值也常常落在 0~100 区间、
# 会通过"看起来像税率"的检测，但重复率明显更低（每个 SKU 数量都不太一样）。
# 用已确认的真实样本校准：真正的税率列 distinct_ratio 落在 0.2~0.5 之间，
# 而误判案例（件数列）是 0.75，取 0.6 作为分界，两边都留出安全余量。
TAX_RATE_MAX_DISTINCT_RATIO = 0.6

# 多 sheet 文件（例如「发票/箱单/合同」三表合一）中，用于识别「箱单」这个 sheet 的关键词
PACKING_LIST_SHEET_HINTS = ["箱单", "装箱单", "packing list", "packing-list", "packinglist"]


def _select_target_sheet(sheet_names: list) -> str:
    """
    从工作簿的所有 sheet 名称中，挑出真正的「箱单」(Packing List) sheet。

    背景：新格式的文件把「发票」「箱单」「合同」三个表格放在同一个 xlsx 的
    三个不同 sheet 里，只有「箱单」这个 sheet 才是逐条列出 SKU/品名/HTS/税率
    的表格，必须精确定位到它，不能再默认读第一个 sheet（第一个 sheet 很可能
    是「发票」，表头列名和数据结构都不一样，会导致解析失败或读出错误数据）。

    旧格式的文件通常只有一个 sheet，这里找不到匹配的名字时会自动退回
    第一个 sheet，行为和以前完全一致，不影响旧格式。
    """
    for name in sheet_names:
        name_norm = str(name).strip().lower()
        for hint in PACKING_LIST_SHEET_HINTS:
            if hint.lower() in name_norm:
                return name
    return sheet_names[0]


def _find_header_row(raw_df: pd.DataFrame) -> int | None:
    """
    扫描原始 DataFrame（无表头读入），找到表头行索引。

    优先找同时包含 SKU / Description / HTS 关键词的行；但有些模板里 SKU
    这一列的表头是空白的（不像"Item Description"/"HTS"那样写了文字），
    这种情况下退而求其次，只要同时有 Description / HTS 关键词就先记下来，
    整个文件扫完都没找到"三者都有"的行时，就用这个退而求其次的候选行
    （SKU 列会在后面用数据特征去定位，见 _locate_sku_column）。
    """
    max_scan_rows = min(50, len(raw_df))
    relaxed_candidate = None
    for row_idx in range(max_scan_rows):
        row_values = [str(v).lower() for v in raw_df.iloc[row_idx].tolist()]
        row_text = " | ".join(row_values)

        has_sku = any(kw in row_text for kw in COLUMN_KEYWORDS["sku"])
        has_desc = any(kw in row_text for kw in COLUMN_KEYWORDS["description"])
        has_hts = any(kw in row_text for kw in COLUMN_KEYWORDS["hts"])

        if has_desc and has_hts and has_sku:
            return row_idx
        if has_desc and has_hts and relaxed_candidate is None:
            relaxed_candidate = row_idx
    return relaxed_candidate


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
    """
    把单元格解析成 0~100 的税率数字，解析不了或超出范围返回 None。

    这里要处理两种历史遗留的存储方式：
    1. 单元格本身就是"6.5"这种已经是百分比数值的数字（数字类型或字符串都可能）。
    2. 单元格被 Excel 设置成"百分比"格式，底层实际存的是 0.065 这种小数
       （Excel 界面上显示成 6.50%，但 pandas/openpyxl 读出来的浮点数就是
       0.065）——这种情况必须乘以 100 换算成 6.5，否则会变成离谱的 0.065%。

    区分方法：如果单元格文本里带 "%" 符号，说明是按"百分号前面的数字就是
    百分比数值"这个约定填的（例如字符串 "6.5%"），不需要再乘 100；
    没带 "%" 符号、且解析出的数值在 0~1 之间（不含 1），大概率是 Excel 百分比
    格式单元格底层的小数，需要乘以 100 换算成真正的百分比数值。
    """
    if pd.isna(val):
        return None
    has_percent_sign = isinstance(val, str) and "%" in val
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
    if not has_percent_sign and 0 < f < 1:
        f = f * 100
    if 0 <= f <= 100:
        return round(f, 4)
    return None


def _column_looks_like_tax_rate(series: pd.Series) -> bool:
    """判断一列是否"看起来像税率"：多数非空值是 0~100 的数字。"""
    non_null = series.dropna()
    non_null = non_null[non_null.astype(str).str.strip() != ""]
    if len(non_null) == 0:
        return False
    valid = sum(1 for v in non_null if _parse_tax_value(v) is not None)
    return (valid / len(non_null)) >= TAX_RATE_DETECT_MIN_RATIO


def _distinct_value_ratio(series: pd.Series) -> float:
    """
    非空值里"不重复值"的占比，越低说明这一列的取值越容易重复。
    用来把"税率"（取值范围有限、大量重复）和"税额/税金"（几乎每行都不同的
    连续型金额）区分开——两者往往都落在 0~100 区间，单看数值范围分不出来。
    """
    non_null = series.dropna()
    non_null = non_null[non_null.astype(str).str.strip() != ""]
    if len(non_null) == 0:
        return 1.0
    return non_null.nunique() / len(non_null)


def _locate_tax_rate_column(raw: pd.DataFrame, header_idx: int, sku_col_idx: int,
                             already_mapped_cols: set) -> int | None:
    """
    定位税率列。策略：收集 SKU 列右边所有"看起来像税率"的候选列，
    取"不重复值占比最低"（最容易重复）的一列；占比相同时取最靠右的一列。

    这么做的原因：这类模板里 SKU 列右边常常同时有好几列数值——「本 SKU 数量」
    「单件重量」「税率」「税额/税金」等，其中不少也恰好落在 0~100 区间、
    会被误判成"像税率"。真正的税率本质上是按 HTS 分类挂钩的，取值范围很
    有限，同一批货的不同 SKU 之间会大量重复；而"税额/税金""每 SKU 件数/箱数"
    这些是数量×单价×税率算出来的连续型金额或逐件计数，几乎每行都不一样。
    所以取重复率最高（不重复值占比最低）的候选列，比单纯"取最靠右一列"更
    可靠；而且额外要求这个占比不能超过 TAX_RATE_MAX_DISTINCT_RATIO —— 如果
    连重复率最高的候选列都达不到"像税率"该有的重复程度，说明这份文件里
    很可能根本没有真正的税率列（很可能是件数/箱数这类列恰好数值也落在
    0~100 区间，被误当成候选），这种情况下宁可整列留空、提示"未找到税率
    列"，也不要把明显不像税率的数据错误地当成税率写进数据库。
    都找不到候选列时返回 None（调用方留空，不报警）。
    """
    data_rows = raw.iloc[header_idx + 1:]
    total_cols = raw.shape[1]

    candidates = []
    for col_idx in range(sku_col_idx + 1, total_cols):
        if col_idx in already_mapped_cols:
            continue
        col = data_rows[col_idx]
        if not _column_looks_like_tax_rate(col):
            continue
        candidates.append((_distinct_value_ratio(col), col_idx))

    # 只保留"重复率"达标的候选（占比不超过上限），排除掉像件数/箱数这种
    # 虽然落在 0~100 区间、但每行都不太一样的列。
    candidates = [c for c in candidates if c[0] <= TAX_RATE_MAX_DISTINCT_RATIO]

    if not candidates:
        return None
    # 不重复值占比最低的优先；占比相同时取列索引最大的（最靠右）
    candidates.sort(key=lambda pair: (pair[0], -pair[1]))
    return candidates[0][1]


def _looks_purely_numeric(val) -> bool:
    """判断一个单元格值是否是"纯数字"（int/float，或能直接转成 float 的字符串）。"""
    if pd.isna(val):
        return False
    if isinstance(val, (int, float)):
        return True
    try:
        float(str(val).strip())
        return True
    except ValueError:
        return False


def _locate_sku_column(raw: pd.DataFrame, header_idx: int, hts_col_idx: int,
                        already_mapped_cols: set) -> int | None:
    """
    定位 SKU 列（用于表头文字里没有写"SKU"的模板）。

    策略：SKU 在这类模板里几乎每一行都有值——哪怕这一行的品名/HTS 是靠
    forward-fill 从上一行继承的（同一产品多个 SKU 只在第一行写品名/HTS），
    SKU 本身仍然逐行填写；而且 SKU 的值基本不是纯数字（都是字母+数字混合
    的编码）。在 HTS 列右边找同时满足"非空率高"和"基本不是纯数字"这两个
    条件的第一列。找不到时返回 None（调用方会报"缺少必要列"错误）。
    """
    data_rows = raw.iloc[header_idx + 1:]
    total_cols = raw.shape[1]

    for col_idx in range(hts_col_idx + 1, total_cols):
        if col_idx in already_mapped_cols:
            continue
        col = data_rows[col_idx]
        non_null = col.dropna()
        non_null = non_null[non_null.astype(str).str.strip() != ""]
        if len(non_null) == 0 or len(col) == 0:
            continue
        fill_ratio = len(non_null) / len(col)
        textual_ratio = (~non_null.apply(_looks_purely_numeric)).mean()
        if fill_ratio >= 0.7 and textual_ratio >= 0.8:
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
        xls = pd.ExcelFile(file_obj_or_path)
        target_sheet = _select_target_sheet(xls.sheet_names)
        raw = xls.parse(sheet_name=target_sheet, header=None)
    except Exception as e:
        return {"success": False, "error": f"无法读取 Excel 文件：{e}", "po_number": None,
                "records": pd.DataFrame(), "warnings": warnings}

    if len(xls.sheet_names) > 1:
        warnings.append(f"该文件包含多个 sheet（{', '.join(xls.sheet_names)}），已自动定位并只解析「{target_sheet}」。")

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

    # SKU 列表头有时是空白的（没写"SKU"这几个字），文字关键词匹配不到时，
    # 退而用数据特征（非空率高 + 基本不是纯数字）在 HTS 列右边找。
    if "sku" not in col_map.values() and "hts" in col_map.values():
        hts_col_idx = [c for c, f in col_map.items() if f == "hts"][0]
        sku_col_idx = _locate_sku_column(raw, header_idx, hts_col_idx, set(col_map.keys()))
        if sku_col_idx is not None:
            col_map[sku_col_idx] = "sku"
            warnings.append("SKU 列没有文字表头，已根据数据特征自动定位。")

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
