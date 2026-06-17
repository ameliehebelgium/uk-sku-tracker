"""
🇬🇧 英国 (UK) 进口 SKU 主数据比对与预警系统
功能1：上传新 PL，自动比对 SKU 的品名/HTS 是否与 UK 主数据库一致，差异预警
功能2：维护 UK SKU 主数据库（新 SKU 自动入库，变更需人工确认后才写入）

⚠️ 本系统专用于英国进口数据，与欧盟 (EU) Risk App 完全独立部署，
   使用独立的 Google Spreadsheet 文件，请勿混用 EU 的 Sheet ID。
"""

import streamlit as st
import pandas as pd

from pl_parser import parse_packing_list
from compare_engine import compare_batch, summarize_comparison, DEFAULT_FUZZY_THRESHOLD
from sheets_db import (
    load_sku_master, load_change_log,
    append_new_skus, update_sku_record, append_change_log,
)

st.set_page_config(page_title="UK SKU 比对预警系统", page_icon="🇬🇧", layout="wide")

# ---------------------------------------------------------------------------
# Session state 初始化
# ---------------------------------------------------------------------------
if "comparison_df" not in st.session_state:
    st.session_state.comparison_df = None
if "source_po" not in st.session_state:
    st.session_state.source_po = None
if "resolved" not in st.session_state:
    st.session_state.resolved = {}  # sku -> "updated" / "ignored"

st.title("🇬🇧 英国进口 SKU 主数据比对与预警系统")
st.caption("本系统专用于英国 (UK) 进口商品数据，独立于欧盟 (EU) Risk App。")

tab_check, tab_database, tab_log = st.tabs(["🔍 上传比对", "🗄️ UK 主数据库", "📋 UK 变更日志"])

# ---------------------------------------------------------------------------
# Tab 1: 上传比对
# ---------------------------------------------------------------------------
with tab_check:
    st.subheader("上传新的 UK Packing List")

    col_upload, col_threshold = st.columns([3, 1])
    with col_upload:
        uploaded_file = st.file_uploader("选择 Excel 文件 (.xlsx)", type=["xlsx"])
    with col_threshold:
        fuzzy_threshold = st.slider(
            "品名相似度阈值", min_value=50, max_value=100,
            value=DEFAULT_FUZZY_THRESHOLD, step=1,
            help="低于此分数视为品名发生实质性变化（满分100为完全一致）。"
        )

    if uploaded_file is not None:
        if st.button("开始解析与比对", type="primary"):
            with st.spinner("解析文件中..."):
                parse_result = parse_packing_list(uploaded_file)

            if not parse_result["success"]:
                st.error(f"解析失败：{parse_result['error']}")
            else:
                for w in parse_result["warnings"]:
                    st.warning(w)

                with st.spinner("正在比对 UK 主数据库..."):
                    master_df = load_sku_master()
                    comparison_df = compare_batch(
                        parse_result["records"], master_df, fuzzy_threshold
                    )

                st.session_state.comparison_df = comparison_df
                st.session_state.source_po = parse_result["po_number"]
                st.session_state.resolved = {}
                st.success(f"解析完成，来源 PO/箱单号：{parse_result['po_number'] or '未识别'}")

    if st.session_state.comparison_df is not None:
        comparison_df = st.session_state.comparison_df
        summary = summarize_comparison(comparison_df)

        st.markdown("---")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("总 SKU 数", summary["total"])
        c2.metric("✅ 一致", summary["match"])
        c3.metric("🔴 HTS 不一致", summary["hts_mismatch"])
        c4.metric("🟡 品名不一致", summary["desc_mismatch"])

        new_df = comparison_df[comparison_df["status"] == "NEW"]
        if not new_df.empty:
            st.markdown("---")
            st.markdown(f"### 🆕 新 SKU（{len(new_df)} 个）— 将自动写入 UK 数据库")
            st.dataframe(
                new_df[["sku", "new_description", "new_hts"]].rename(columns={
                    "new_description": "品名", "new_hts": "HTS"
                }),
                use_container_width=True, hide_index=True,
            )
            if st.button("✅ 确认写入这些新 SKU 到 UK 数据库"):
                append_new_skus(
                    new_df.rename(columns={"new_description": "description", "new_hts": "hts"}),
                    st.session_state.source_po,
                )
                st.success(f"已写入 {len(new_df)} 个新 SKU。")
                st.cache_resource.clear()

        mismatch_df = comparison_df[comparison_df["status"].isin(["HTS_MISMATCH", "DESC_MISMATCH"])]
        if not mismatch_df.empty:
            st.markdown("---")
            st.markdown(f"### ⚠️ 差异预警（{len(mismatch_df)} 个）— 请逐条确认处理方式")

            for _, row in mismatch_df.iterrows():
                sku = row["sku"]
                already_resolved = st.session_state.resolved.get(sku)

                icon = "🔴" if row["status"] == "HTS_MISMATCH" else "🟡"
                label = "HTS 不一致" if row["status"] == "HTS_MISMATCH" else f"品名不一致（相似度 {row['desc_similarity']}）"

                with st.container(border=True):
                    st.markdown(f"{icon} **{sku}** — {label}")
                    col_old, col_new = st.columns(2)
                    with col_old:
                        st.caption("UK 数据库现有记录")
                        st.write(f"品名：{row['old_description']}")
                        st.write(f"HTS：{row['old_hts']}")
                    with col_new:
                        st.caption("本次 PL 记录")
                        st.write(f"品名：{row['new_description']}")
                        st.write(f"HTS：{row['new_hts']}")

                    if already_resolved:
                        st.info(f"已处理：{ '已更新数据库' if already_resolved == 'updated' else '已忽略，保留原记录' }")
                    else:
                        btn_col1, btn_col2 = st.columns(2)
                        with btn_col1:
                            if st.button(f"✅ 用新值更新数据库", key=f"update_{sku}"):
                                update_sku_record(sku, row["new_description"], row["new_hts"], st.session_state.source_po)
                                log_entries = []
                                if row["status"] == "HTS_MISMATCH":
                                    log_entries.append({
                                        "sku": sku, "field_changed": "hts",
                                        "old_value": row["old_hts"], "new_value": row["new_hts"],
                                        "source_po": st.session_state.source_po, "resolution": "updated",
                                    })
                                if row["status"] == "DESC_MISMATCH" or row["status"] == "HTS_MISMATCH":
                                    log_entries.append({
                                        "sku": sku, "field_changed": "description",
                                        "old_value": row["old_description"], "new_value": row["new_description"],
                                        "source_po": st.session_state.source_po, "resolution": "updated",
                                    })
                                append_change_log(log_entries)
                                st.session_state.resolved[sku] = "updated"
                                st.cache_resource.clear()
                                st.rerun()
                        with btn_col2:
                            if st.button(f"❌ 忽略，保留原记录", key=f"ignore_{sku}"):
                                field = "hts" if row["status"] == "HTS_MISMATCH" else "description"
                                append_change_log([{
                                    "sku": sku, "field_changed": field,
                                    "old_value": row["old_hts"] if field == "hts" else row["old_description"],
                                    "new_value": row["new_hts"] if field == "hts" else row["new_description"],
                                    "source_po": st.session_state.source_po, "resolution": "ignored",
                                }])
                                st.session_state.resolved[sku] = "ignored"
                                st.rerun()

        if summary["hts_mismatch"] == 0 and summary["desc_mismatch"] == 0 and summary["new"] == 0:
            st.info("所有 SKU 均与 UK 数据库一致，无需处理。")

# ---------------------------------------------------------------------------
# Tab 2: 主数据库浏览
# ---------------------------------------------------------------------------
with tab_database:
    st.subheader("UK SKU 主数据库")
    master_df = load_sku_master()
    search = st.text_input("搜索 SKU 或品名关键词")
    display_df = master_df
    if search:
        mask = (
            master_df["sku"].str.contains(search, case=False, na=False)
            | master_df["description"].str.contains(search, case=False, na=False)
        )
        display_df = master_df[mask]
    st.caption(f"共 {len(master_df)} 条记录，当前显示 {len(display_df)} 条")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Tab 3: 变更日志
# ---------------------------------------------------------------------------
with tab_log:
    st.subheader("UK 变更审计日志")
    st.caption("记录每一次检测到的品名/HTS 差异及处理结果，作为合规追溯依据。")
    log_df = load_change_log()
    if log_df.empty:
        st.info("暂无记录。")
    else:
        st.dataframe(log_df.sort_values("timestamp", ascending=False), use_container_width=True, hide_index=True)
