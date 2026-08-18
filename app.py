"""
🇬🇧 英国 (UK) 进口 SKU 主数据比对与预警系统
功能1：登录后，上传新 PL（支持多个 xlsx 或一个 zip），自动比对品名/HTS/税率是否与 UK 主数据库一致，差异预警
功能2：维护 UK SKU 主数据库（新 SKU 自动入库，变更需人工确认后才写入），可随时导出全库
功能3：每次上传比对后，可下载本次比对结果汇总

⚠️ 本系统专用于英国进口数据，与欧盟 (EU) Risk App 完全独立部署，
   使用独立的 Google Spreadsheet 文件，请勿混用 EU 的 Sheet ID。
"""

import io
import zipfile

import streamlit as st
import pandas as pd

from pl_parser import parse_packing_list
from compare_engine import compare_batch, summarize_comparison, DEFAULT_FUZZY_THRESHOLD
from sheets_db import (
    load_sku_master, load_change_log,
    append_new_skus, update_sku_record, append_change_log, bulk_update_tax_rates,
)

st.set_page_config(page_title="UK SKU 比对预警系统", page_icon="🇬🇧", layout="wide")


# ---------------------------------------------------------------------------
# 登录校验
# 用户名/密码优先从 st.secrets 读取（部署环境改密码更方便），读不到就用默认值兜底。
# 这是一个简单的入口门禁，不是完整的账号体系（没有多用户、没有权限分级）。
# ---------------------------------------------------------------------------
DEFAULT_ADMIN_USERNAME = "Admin"
DEFAULT_ADMIN_PASSWORD = "Admin123"


def _check_login() -> bool:
    if st.session_state.get("logged_in"):
        return True

    st.title("🇬🇧 UK SKU 比对预警系统")
    st.caption("请登录后使用")

    with st.form("login_form"):
        username = st.text_input("用户名")
        password = st.text_input("密码", type="password")
        submitted = st.form_submit_button("登录", type="primary")

    if submitted:
        admin_user = st.secrets.get("ADMIN_USERNAME", DEFAULT_ADMIN_USERNAME)
        admin_pass = st.secrets.get("ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD)
        if username == admin_user and password == admin_pass:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("用户名或密码错误")

    return False


if not _check_login():
    st.stop()


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def _extract_excel_files(uploaded_files):
    """
    把用户拖入的文件展开成 [(文件名, 文件对象), ...]：
    - .xlsx 直接收录
    - .zip 解压后，取里面所有 .xlsx（跳过 __MACOSX 等垃圾目录）
    - 其他类型的文件跳过，并记录提示
    """
    excel_files = []
    notes = []
    for uf in uploaded_files:
        name_lower = uf.name.lower()
        if name_lower.endswith(".zip"):
            try:
                with zipfile.ZipFile(uf) as zf:
                    found_any = False
                    for member in zf.namelist():
                        if member.startswith("__MACOSX") or member.endswith("/"):
                            continue
                        if member.lower().endswith(".xlsx"):
                            excel_files.append((f"{uf.name} / {member}", io.BytesIO(zf.read(member))))
                            found_any = True
                        else:
                            notes.append(f"{uf.name} 内的「{member}」不是 xlsx，已跳过")
                    if not found_any:
                        notes.append(f"{uf.name} 里没有找到任何 xlsx 文件")
            except zipfile.BadZipFile:
                notes.append(f"{uf.name} 不是有效的 zip 文件，已跳过")
        elif name_lower.endswith(".xlsx"):
            excel_files.append((uf.name, uf))
        else:
            notes.append(f"{uf.name} 既不是 xlsx 也不是 zip，已跳过")
    return excel_files, notes


def _df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()


def _tax_proposed_change(row) -> bool:
    """
    本次 PL 记录的税率是否跟数据库原值不一样（用于判断 HTS/品名不一致的卡片
    要不要额外弹出"税率单独确认"的选项）。

    compare_engine 里 new_tax_rate 已经处理过"本次没读到税率就退回用原值"的
    情况（effective_new_tax_rate），所以这里只需要简单比较字符串——如果本次
    没有新税率，new_tax_rate 早就等于 old_tax_rate 了，不会误触发。
    """
    new_val = str(row.get("new_tax_rate") or "").strip()
    old_val = str(row.get("old_tax_rate") or "").strip()
    return new_val != "" and new_val != old_val


def _combined_comparison_df(file_results: dict) -> pd.DataFrame:
    """把本次上传的所有文件的比对结果合并成一个 DataFrame，加一列标注来源文件。"""
    frames = []
    for fname, info in file_results.items():
        cdf = info.get("comparison_df")
        if cdf is not None and not cdf.empty:
            cdf = cdf.copy()
            cdf.insert(0, "source_file", fname)
            cdf.insert(1, "source_po", info["parse_result"].get("po_number") or "")
            frames.append(cdf)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


# ---------------------------------------------------------------------------
# Session state 初始化
# ---------------------------------------------------------------------------
if "file_results" not in st.session_state:
    st.session_state.file_results = {}  # 文件名 -> {"parse_result": ..., "comparison_df": ...}
if "resolved" not in st.session_state:
    st.session_state.resolved = {}  # (文件名, sku) -> "updated" / "ignored"

st.title("🇬🇧 英国进口 SKU 主数据比对与预警系统")
st.caption("本系统专用于英国 (UK) 进口商品数据，独立于欧盟 (EU) Risk App。")

tab_check, tab_database, tab_log = st.tabs(["🔍 上传比对", "🗄️ UK 主数据库", "📋 UK 变更日志"])

# ---------------------------------------------------------------------------
# Tab 1: 上传比对
# ---------------------------------------------------------------------------
with tab_check:
    st.subheader("上传新的 UK Packing List")
    st.caption("支持拖拽单个/多个 Excel (.xlsx) 文件，或者一个包含多个 xlsx 的 ZIP 压缩包。")

    col_upload, col_threshold = st.columns([3, 1])
    with col_upload:
        uploaded_files = st.file_uploader(
            "选择或拖拽文件", type=["xlsx", "zip"], accept_multiple_files=True
        )
    with col_threshold:
        fuzzy_threshold = st.slider(
            "品名相似度阈值", min_value=50, max_value=100,
            value=DEFAULT_FUZZY_THRESHOLD, step=1,
            help="低于此分数视为品名发生实质性变化（满分100为完全一致）。"
        )

    if uploaded_files:
        if st.button("开始解析与比对", type="primary"):
            excel_files, notes = _extract_excel_files(uploaded_files)
            for n in notes:
                st.warning(n)

            if not excel_files:
                st.error("没有可解析的 xlsx 文件。")
            else:
                master_df = load_sku_master()
                file_results = {}
                total_autofilled = 0
                with st.spinner(f"正在解析并比对 {len(excel_files)} 个文件..."):
                    for fname, fobj in excel_files:
                        parse_result = parse_packing_list(fobj)
                        comparison_df = None
                        if parse_result["success"]:
                            comparison_df = compare_batch(
                                parse_result["records"], master_df, fuzzy_threshold
                            )
                            # 税率从空值补全为具体数值：品名、HTS 都没变，不是冲突，
                            # 不需要人工逐条确认，这里直接写入数据库。只在"开始解析与比对"
                            # 这次点击里跑一次，不会随其他按钮的页面刷新重复写入。
                            # 用批量写入（bulk_update_tax_rates），而不是对每个 SKU 单独
                            # 调用 update_sku_record——后者一个 SKU 就要发好几个 Google
                            # Sheets API 请求，一批几十个 SKU 循环调用很容易在几秒内打满
                            # 每分钟请求配额、触发 429 报错。
                            autofill_df = comparison_df[comparison_df["status"] == "TAX_AUTOFILL"]
                            if not autofill_df.empty:
                                bulk_update_tax_rates([
                                    {"sku": row["sku"], "tax_rate": row["new_tax_rate"]}
                                    for _, row in autofill_df.iterrows()
                                ])
                                log_entries = [{
                                    "sku": row["sku"], "field_changed": "tax_rate",
                                    "old_value": row["old_tax_rate"], "new_value": row["new_tax_rate"],
                                    "source_po": parse_result["po_number"], "resolution": "auto_filled",
                                } for _, row in autofill_df.iterrows()]
                                append_change_log(log_entries)
                                total_autofilled += len(autofill_df)
                        file_results[fname] = {
                            "parse_result": parse_result,
                            "comparison_df": comparison_df,
                        }

                st.session_state.file_results = file_results
                st.session_state.resolved = {}
                st.success(f"已解析 {len(excel_files)} 个文件。")
                if total_autofilled > 0:
                    st.success(f"✅ {total_autofilled} 个 SKU 的税率从空值自动补全为本次 PL 记录的数值，已直接写入 UK 数据库。")

    file_results = st.session_state.file_results

    if file_results:
        # ---- 本次结果汇总下载（合并所有文件） ----
        combined_df = _combined_comparison_df(file_results)
        if not combined_df.empty:
            st.markdown("---")
            dl_col1, dl_col2 = st.columns([1, 3])
            with dl_col1:
                st.download_button(
                    "⬇️ 下载本次结果汇总",
                    data=_df_to_excel_bytes(combined_df, sheet_name="本次比对结果"),
                    file_name="UK_本次比对结果汇总.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with dl_col2:
                st.caption("包含本次上传的所有文件的比对结果（新SKU/一致/HTS不一致/品名不一致），每行标注来源文件与PO。")

        # ---- 逐文件展示 ----
        for fname, info in file_results.items():
            parse_result = info["parse_result"]
            comparison_df = info["comparison_df"]

            st.markdown("---")
            st.markdown(f"## 📄 {fname}")

            if not parse_result["success"]:
                st.error(f"解析失败：{parse_result['error']}")
                continue

            for w in parse_result["warnings"]:
                st.warning(w)

            st.caption(f"来源 PO/箱单号：{parse_result['po_number'] or '未识别'}")

            summary = summarize_comparison(comparison_df)
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("总 SKU 数", summary["total"])
            c2.metric("✅ 一致", summary["match"])
            c3.metric("🔴 HTS 不一致", summary["hts_mismatch"])
            c4.metric("🟡 品名不一致", summary["desc_mismatch"])
            c5.metric("🔵 税率不一致", summary["tax_mismatch"])
            c6.metric("🟢 税率自动补全", summary["tax_autofill"])

            new_df = comparison_df[comparison_df["status"] == "NEW"]
            if not new_df.empty:
                st.markdown(f"#### 🆕 新 SKU（{len(new_df)} 个）— 将自动写入 UK 数据库")
                st.dataframe(
                    new_df[["sku", "new_description", "new_hts", "new_tax_rate"]].rename(columns={
                        "new_description": "品名", "new_hts": "HTS", "new_tax_rate": "税率"
                    }),
                    width='stretch', hide_index=True,
                )
                if st.button("✅ 确认写入这些新 SKU 到 UK 数据库", key=f"write_new_{fname}"):
                    append_new_skus(
                        new_df.rename(columns={
                            "new_description": "description", "new_hts": "hts", "new_tax_rate": "tax_rate",
                        }),
                        parse_result["po_number"],
                    )
                    st.success(f"已写入 {len(new_df)} 个新 SKU。")

            mismatch_df = comparison_df[comparison_df["status"].isin(
                ["HTS_MISMATCH", "DESC_MISMATCH", "TAX_MISMATCH"]
            )]
            if not mismatch_df.empty:
                st.markdown(f"#### ⚠️ 差异预警（{len(mismatch_df)} 个）— 请逐条确认处理方式")

                for _, row in mismatch_df.iterrows():
                    sku = row["sku"]
                    resolve_key = (fname, sku)
                    already_resolved = st.session_state.resolved.get(resolve_key)

                    if row["status"] == "HTS_MISMATCH":
                        icon, label = "🔴", "HTS 不一致"
                    elif row["status"] == "DESC_MISMATCH":
                        icon, label = "🟡", f"品名不一致（相似度 {row['desc_similarity']}）"
                    else:
                        icon, label = "🔵", f"税率不一致（{row['old_tax_rate'] or '（空）'} → {row['new_tax_rate']}）"

                    with st.container(border=True):
                        st.markdown(f"{icon} **{sku}** — {label}")
                        col_old, col_new = st.columns(2)
                        with col_old:
                            st.caption("UK 数据库现有记录")
                            st.write(f"品名：{row['old_description']}")
                            st.write(f"HTS：{row['old_hts']}")
                            st.write(f"税率：{row['old_tax_rate'] or '（空）'}")
                        with col_new:
                            st.caption("本次 PL 记录")
                            st.write(f"品名：{row['new_description']}")
                            st.write(f"HTS：{row['new_hts']}")
                            st.write(f"税率：{row['new_tax_rate'] or '（空）'}")

                        if already_resolved:
                            st.info(f"已处理：{'已更新数据库' if already_resolved == 'updated' else '已忽略，保留原记录'}")
                        else:
                            # HTS/品名不一致的同时，本次 PL 读到的税率如果跟数据库原值也不一样，
                            # 这是两件独立的事——HTS/品名的差异可能是真实变更，但同一份 PL 上
                            # 的税率完全可能是数据源本身的错误（跟 HTS/品名是否可信没有必然关系）。
                            # 所以这里单独给一个税率处理方式的选项，不强制跟 HTS/品名绑在一起。
                            # 默认选中"采用新税率"，跟以前的行为保持一致，用户判断这次税率不可信
                            # 时可以手动切换成"保留原税率"，两者互不影响地一起点击下面的确认按钮。
                            show_tax_choice = (
                                row["status"] in ("HTS_MISMATCH", "DESC_MISMATCH")
                                and _tax_proposed_change(row)
                            )
                            tax_accepted = True
                            if show_tax_choice:
                                st.caption(
                                    f"⚠️ 本次 PL 记录的税率（{row['new_tax_rate']}）与数据库原值"
                                    f"（{row['old_tax_rate'] or '（空）'}）也不一样，这跟上面的"
                                    f"{'HTS' if row['status'] == 'HTS_MISMATCH' else '品名'}差异是两件独立的事，"
                                    f"请单独确认税率是否可信："
                                )
                                tax_decision = st.radio(
                                    "税率处理方式",
                                    options=[
                                        f"采用本次 PL 的新税率（{row['new_tax_rate']}）",
                                        f"保留数据库原税率（{row['old_tax_rate'] or '（空）'}）",
                                    ],
                                    key=f"taxchoice_{fname}_{sku}",
                                    horizontal=True,
                                    label_visibility="collapsed",
                                )
                                tax_accepted = tax_decision.startswith("采用本次")

                            btn_col1, btn_col2 = st.columns(2)
                            with btn_col1:
                                if st.button("✅ 用新值更新数据库", key=f"update_{fname}_{sku}"):
                                    effective_tax_rate = row["new_tax_rate"] if tax_accepted else row["old_tax_rate"]
                                    update_sku_record(
                                        sku, row["new_description"], row["new_hts"],
                                        parse_result["po_number"], new_tax_rate=effective_tax_rate,
                                    )
                                    log_entries = []
                                    if row["status"] == "HTS_MISMATCH":
                                        log_entries.append({
                                            "sku": sku, "field_changed": "hts",
                                            "old_value": row["old_hts"], "new_value": row["new_hts"],
                                            "source_po": parse_result["po_number"], "resolution": "updated",
                                        })
                                    if row["status"] in ("DESC_MISMATCH", "HTS_MISMATCH"):
                                        log_entries.append({
                                            "sku": sku, "field_changed": "description",
                                            "old_value": row["old_description"], "new_value": row["new_description"],
                                            "source_po": parse_result["po_number"], "resolution": "updated",
                                        })
                                    if show_tax_choice and not tax_accepted:
                                        # 税率被单独拒绝了：记下"本次 PL 提出过这个税率，但被人工
                                        # 判定不可信、没有采纳"，保留完整审计轨迹，跟 HTS/品名的
                                        # "updated" 结果分开记录。
                                        log_entries.append({
                                            "sku": sku, "field_changed": "tax_rate",
                                            "old_value": row["old_tax_rate"], "new_value": row["new_tax_rate"],
                                            "source_po": parse_result["po_number"], "resolution": "rejected_kept_old",
                                        })
                                    elif effective_tax_rate != row["old_tax_rate"]:
                                        log_entries.append({
                                            "sku": sku, "field_changed": "tax_rate",
                                            "old_value": row["old_tax_rate"], "new_value": effective_tax_rate,
                                            "source_po": parse_result["po_number"], "resolution": "updated",
                                        })
                                    append_change_log(log_entries)
                                    st.session_state.resolved[resolve_key] = "updated"
                                    st.rerun()
                            with btn_col2:
                                if st.button("❌ 忽略，保留原记录", key=f"ignore_{fname}_{sku}"):
                                    if row["status"] == "HTS_MISMATCH":
                                        field, old_val, new_val = "hts", row["old_hts"], row["new_hts"]
                                    elif row["status"] == "DESC_MISMATCH":
                                        field, old_val, new_val = "description", row["old_description"], row["new_description"]
                                    else:
                                        field, old_val, new_val = "tax_rate", row["old_tax_rate"], row["new_tax_rate"]
                                    append_change_log([{
                                        "sku": sku, "field_changed": field,
                                        "old_value": old_val, "new_value": new_val,
                                        "source_po": parse_result["po_number"], "resolution": "ignored",
                                    }])
                                    st.session_state.resolved[resolve_key] = "ignored"
                                    st.rerun()

            if (summary["hts_mismatch"] == 0 and summary["desc_mismatch"] == 0
                    and summary["tax_mismatch"] == 0 and summary["new"] == 0
                    and summary["tax_autofill"] == 0):
                st.info("所有 SKU 均与 UK 数据库一致，无需处理。")

# ---------------------------------------------------------------------------
# Tab 2: 主数据库浏览
# ---------------------------------------------------------------------------
with tab_database:
    st.subheader("UK SKU 主数据库")
    master_df = load_sku_master()

    col_search, col_export = st.columns([3, 1])
    with col_search:
        search = st.text_input("搜索 SKU 或品名关键词")
    with col_export:
        st.write("")
        st.download_button(
            "⬇️ 导出全部主数据库",
            data=_df_to_excel_bytes(master_df, sheet_name="UK_SKU_Master"),
            file_name="UK_SKU_Master_全库导出.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    display_df = master_df
    if search:
        mask = (
            master_df["sku"].str.contains(search, case=False, na=False)
            | master_df["description"].str.contains(search, case=False, na=False)
        )
        display_df = master_df[mask]
    st.caption(f"共 {len(master_df)} 条记录，当前显示 {len(display_df)} 条")
    st.dataframe(display_df, width='stretch', hide_index=True)

# ---------------------------------------------------------------------------
# Tab 3: 变更日志
# ---------------------------------------------------------------------------
with tab_log:
    st.subheader("UK 变更审计日志")
    st.caption("记录每一次检测到的品名/HTS/税率差异及处理结果，作为合规追溯依据。")
    log_df = load_change_log()
    if log_df.empty:
        st.info("暂无记录。")
    else:
        st.dataframe(log_df.sort_values("timestamp", ascending=False), width='stretch', hide_index=True)
