"""
SKU 比对引擎
负责把新解析出的 PL 记录，跟 SKU_Master 主数据库逐条比对，
判定每个 SKU 属于以下哪种状态：

    NEW          - 数据库里没有这个 SKU（全新 SKU，待写入）
    MATCH        - 品名（模糊匹配通过）+ HTS + 税率均一致（忽略，不预警）
    HTS_MISMATCH - HTS 不一致（无论品名/税率是否一致）—— 🔴 红色，最高优先级
    DESC_MISMATCH- HTS 一致，但品名模糊匹配低于阈值 —— 🟡 黄色，次高优先级
    TAX_MISMATCH - 品名、HTS 都一致，数据库里原本也有税率，但跟本次 PL 读到的
                   不一样 —— 🔵 蓝色，真正的冲突，需要人工确认
    TAX_AUTOFILL - 品名、HTS 都一致，数据库里原本这个 SKU 的税率是空的，本次
                   PL 读到了具体数值 —— 不是"冲突"，是补全缺失数据，不需要人工
                   确认，由调用方直接写入数据库

设计原则：
- HTS 比对是精确匹配（哪怕一位数字不同也要报警，海关编码不允许任何容忍度）。
- 品名比对是模糊匹配（默认阈值 85，基于 rapidfuzz token_sort_ratio，
  能容忍大小写、多余空格、词序轻微差异，但仍能抓出实质性的描述变化）。
- 阈值在外部可调，方便后续根据实际误报情况调整。
- 税率比对：只有当本次 PL 确实解析出了税率（new_tax_rate 非空）才参与判定，
  没解析到税率的文件不会因为"税率留空"而触发误报。
  税率不一致又分两种情况，处理方式不同：
    1. 数据库里原本就有税率、这次读到的不一样 —— 说明数据可能有分歧
       （比如税则归类变了，或者之前录错了），必须交给人工确认才写入。
    2. 数据库里原本是空的、这次读到了具体数值 —— 只是把缺失的数据补上，
       不存在"哪个对哪个错"的冲突，不需要人工确认，可以直接写入。
"""

import pandas as pd
from rapidfuzz import fuzz


DEFAULT_FUZZY_THRESHOLD = 85  # 0-100，越高越严格


def _normalize_desc(text: str) -> str:
    """品名比对前的归一化：大小写、首尾空格、多余空格。"""
    if pd.isna(text):
        return ""
    return " ".join(str(text).strip().upper().split())


def _normalize_hts(text: str) -> str:
    """HTS 比对前的归一化：只去除首尾空格和内部空格/连字符，不改变数字本身。"""
    if pd.isna(text):
        return ""
    return str(text).strip().replace(" ", "").replace("-", "")


def _normalize_tax(text):
    """
    税率比对前的归一化。能解析成数字的（如 "6.5%"、"6.5"、6.5）统一转成浮点数比较，
    这样 "6.5%" 和 "6.50%" 会被视为相同；解析不了的留作字符串兜底比较；
    空值统一返回 None。
    """
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return None
    text = str(text).strip()
    if text == "":
        return None
    try:
        return round(float(text.replace("%", "").replace(",", ".")), 4)
    except ValueError:
        return text.upper()


def compare_sku(new_sku: str, new_desc: str, new_hts: str,
                 master_df: pd.DataFrame,
                 fuzzy_threshold: int = DEFAULT_FUZZY_THRESHOLD,
                 new_tax_rate: str = "") -> dict:
    """
    比对单个 SKU。

    Args:
        master_df: 主数据库 DataFrame，至少包含 sku / description / hts 列
        new_tax_rate: 本次 PL 解析出的税率（已格式化为 "6.5%" 这种字符串，可能为空）。
            税率不参与 NEW/MATCH/MISMATCH 的判定，只是随比对结果一起带出来，
            方便写入主数据库或展示。

    Returns:
        dict，描述比对结果，供上层汇总展示。
    """
    existing = master_df[master_df["sku"] == new_sku]

    if existing.empty:
        return {
            "sku": new_sku,
            "status": "NEW",
            "new_description": new_desc,
            "new_hts": new_hts,
            "new_tax_rate": new_tax_rate,
            "old_description": None,
            "old_hts": None,
            "old_tax_rate": None,
            "desc_similarity": None,
        }

    old_row = existing.iloc[0]
    old_desc = old_row["description"]
    old_hts = old_row["hts"]
    old_tax_rate = old_row["tax_rate"] if "tax_rate" in old_row else ""

    hts_match = _normalize_hts(new_hts) == _normalize_hts(old_hts)
    similarity = fuzz.token_sort_ratio(_normalize_desc(new_desc), _normalize_desc(old_desc))
    desc_match = similarity >= fuzzy_threshold

    # 税率比对：只有本次真的解析到税率（new_tax_rate 非空）才参与判定，
    # 避免"这次文件没有税率列"被误判成"税率被清空了"。
    old_tax_missing = _normalize_tax(old_tax_rate) is None
    tax_changed = bool(new_tax_rate) and _normalize_tax(new_tax_rate) != _normalize_tax(old_tax_rate)

    if not hts_match:
        status = "HTS_MISMATCH"
    elif not desc_match:
        status = "DESC_MISMATCH"
    elif tax_changed:
        # 数据库里原本没有税率 -> 补全缺失数据，不算冲突，不需要人工确认；
        # 数据库里原本有税率、跟本次读到的不一样 -> 真正的冲突，需要人工确认。
        status = "TAX_AUTOFILL" if old_tax_missing else "TAX_MISMATCH"
    else:
        status = "MATCH"

    # 税率留空时不要用空值覆盖数据库里已有的税率（本次没读到税率不代表税率变了）
    effective_new_tax_rate = new_tax_rate if new_tax_rate else old_tax_rate

    return {
        "sku": new_sku,
        "status": status,
        "new_description": new_desc,
        "new_hts": new_hts,
        "new_tax_rate": effective_new_tax_rate,
        "old_description": old_desc,
        "old_hts": old_hts,
        "old_tax_rate": old_tax_rate,
        "desc_similarity": round(similarity, 1),
    }


def compare_batch(new_records: pd.DataFrame, master_df: pd.DataFrame,
                   fuzzy_threshold: int = DEFAULT_FUZZY_THRESHOLD) -> pd.DataFrame:
    """
    批量比对一整份 PL 解析出的记录。

    Args:
        new_records: columns = [sku, description, hts, tax_rate]（tax_rate 可选，没有就当空处理）
        master_df:   columns = [sku, description, hts, tax_rate, ...其他元数据列]

    Returns:
        DataFrame，每行一个比对结果，包含 status / 新旧值 / 相似度
    """
    has_tax_col = "tax_rate" in new_records.columns
    results = []
    for _, row in new_records.iterrows():
        new_tax_rate = row["tax_rate"] if has_tax_col else ""
        result = compare_sku(row["sku"], row["description"], row["hts"], master_df,
                              fuzzy_threshold, new_tax_rate=new_tax_rate)
        results.append(result)

    return pd.DataFrame(results)


def summarize_comparison(comparison_df: pd.DataFrame) -> dict:
    """生成比对结果的统计摘要，用于界面顶部展示。"""
    if comparison_df.empty:
        return {"total": 0, "new": 0, "match": 0, "hts_mismatch": 0, "desc_mismatch": 0,
                "tax_mismatch": 0, "tax_autofill": 0}

    counts = comparison_df["status"].value_counts()
    return {
        "total": len(comparison_df),
        "new": int(counts.get("NEW", 0)),
        "match": int(counts.get("MATCH", 0)),
        "hts_mismatch": int(counts.get("HTS_MISMATCH", 0)),
        "desc_mismatch": int(counts.get("DESC_MISMATCH", 0)),
        "tax_mismatch": int(counts.get("TAX_MISMATCH", 0)),
        "tax_autofill": int(counts.get("TAX_AUTOFILL", 0)),
    }
