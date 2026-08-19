"""
Google Sheets 数据库连接层 —— 英国 (UK) 进口数据专属

⚠️ 本系统专门用于英国进口商品的 SKU/品名/HTS 数据管理，与欧盟 (EU) Risk App
   完全独立：使用独立的 Google Spreadsheet 文件（不同的 SHEET_ID），
   Sheet 名称统一加 UK_ 前缀，避免未来误连接到 EU 数据源或两者混淆。

   两个核心 Sheet：
    1. UK_SKU_Master  - 英国主数据库（当前生效的 SKU / 品名 / HTS）
    2. UK_Change_Log  - 英国变更审计日志（只增不改）

认证方式与 EU Risk App 保持一致（同一个 Google 账号下的 service account
均可复用，只需确保该 service account 被共享到这个新建的 UK Spreadsheet）。
"""

import re
import time
import gspread
import pandas as pd
import streamlit as st
from datetime import datetime
from google.oauth2.service_account import Credentials


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Sheet 名称统一加 UK_ 前缀
SKU_MASTER_SHEET_NAME = "UK_SKU_Master"
CHANGE_LOG_SHEET_NAME = "UK_Change_Log"

SKU_MASTER_HEADERS = [
    "sku", "description", "hts", "tax_rate",
    "first_seen_date", "last_updated_date", "source_po",
]

CHANGE_LOG_HEADERS = [
    "timestamp", "sku", "field_changed", "old_value", "new_value",
    "source_po", "resolution",  # resolution: "updated" / "ignored" / "pending"
]

# 429（配额超限）、500/503（Google 那边偶发的服务端错误）都是"过一会儿再试
# 大概率能成功"的临时性错误，值得重试；其他错误（比如权限不对、表不存在）
# 重试也没用，直接抛出交给上层处理。
_RETRYABLE_STATUS_CODES = {429, 500, 503}
# 429 是"每分钟"级别的配额，重试窗口要盖过 60 秒才有意义——之前 5 次、
# 总共约 22 秒的重试窗口，遇到真的把当分钟配额打满的情况还是不够，
# 这里加到 7 次、总共约 2 分钟，确保能等到配额窗口刷新。
_RETRY_MAX_ATTEMPTS = 7
_RETRY_BASE_DELAY_SECONDS = 2


def _call_with_retry(func, *args, **kwargs):
    """
    包一层重试 + 指数退避，用来发实际的 Google Sheets API 请求。

    背景：这套 App 在测试/使用高峰期，短时间内会有好几个动作都要读写
    Google Sheets（打开表、读主数据库、写变更日志……），很容易撞上 Google
    "每分钟/每100秒请求数"的配额上限，报 429 APIError 直接把整个页面搞崩。
    大多数情况下配额是按滚动窗口算的，等个一两秒重试一次基本就能通过，
    不需要真的让用户看到一个红色报错框。
    """
    last_err = None
    for attempt in range(_RETRY_MAX_ATTEMPTS):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = getattr(e, "code", None)
            if status not in _RETRYABLE_STATUS_CODES or attempt == _RETRY_MAX_ATTEMPTS - 1:
                raise
            last_err = e
            time.sleep(_RETRY_BASE_DELAY_SECONDS * (2 ** attempt))
    raise last_err


@st.cache_resource
def _get_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource
def _get_spreadsheet():
    """
    打开本系统专属的 UK Spreadsheet。
    用 cache_resource 缓存住这个连接对象——client.open_by_key() 本身会消耗一次
    Google Sheets API 的读配额，不缓存的话，每次 Streamlit 重新跑脚本（哪怕只是
    点了个无关的按钮）都会重新打开一次，很容易把「每分钟读请求数」的免费额度打满。

    注意：secrets 中的 UK_SHEET_ID 必须指向一个独立于 EU Risk App 的
    全新 Google Sheet 文件，不能复用 EU 那边的 SHEET_ID。
    """
    client = _get_client()
    sheet_id = st.secrets["UK_SHEET_ID"]
    return _call_with_retry(client.open_by_key, sheet_id)


# ---------------------------------------------------------------------------
# UK_SKU_Master 表头/数据错位问题：自动检测 + 自动修复
#
# 背景（一次真实发生过的事故）：UK_SKU_Master 的表头行是在"税率"字段还没
# 加入 SKU_MASTER_HEADERS 之前就建好的（旧顺序：sku, description, hts,
# first_seen_date, last_updated_date, source_po），后来代码把 tax_rate 插入
# 成第 4 个字段，但表头行从来没有跟着更新过（_get_or_create_worksheet 原来
# 的逻辑只在"整行完全空白"时才会写表头，已存在的表头不会被自动纠正）。
# 结果是写入代码一直按新顺序物理写数据（税率在第 4 列），表头却还是旧顺序，
# 导致用表头名字读数据时 first_seen_date / last_updated_date / source_po
# 全部读错列，税率永远读成空的——而且是静默发生的，直到导出全库看 Excel
# 才发现，此时已经有 1000+ 行数据是这个错位状态。
#
# 下面这组函数在每次"找表"时顺手检查一下表头是否跟代码一致；如果发现是
# UK_SKU_Master 这张表、而且错位的行都能被规则明确识别（不是什么诡异的
# 未知格式），就自动把数据搬回正确的物理列位置、顺手把表头也改对，全程
# 不需要人工介入、不需要额外跑脚本、不需要把 Google 凭据交出去——反正
# 部署在 Streamlit Cloud 上的这个 App 本来就有权限读写这张表。
# 只有遇到规则识别不了的行时才会放弃自动修复，转而报错停止，避免在没把握
# 的情况下瞎改数据；这种情况下可以用 migrate_uk_sku_master.py 手动排查。
# ---------------------------------------------------------------------------
_SKU_MASTER_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_SKU_MASTER_PO_RE = re.compile(r"^[A-Za-z]{2,}-")


def _classify_legacy_sku_master_row(row: list) -> str:
    """
    判断 UK_SKU_Master 里一行数据当前是"新物理布局"（已经跟 SKU_MASTER_HEADERS
    一致，不用动）还是"旧物理布局"（税率功能上线前写入的，D/E 是日期、F 是
    来源 PO、G/H 为空，需要整体右移一列）。跟 migrate_uk_sku_master.py 里的
    判定规则保持一致。
    """
    g = (row[6] if len(row) > 6 else "").strip()
    if g:
        return "new"
    d = (row[3] if len(row) > 3 else "").strip()
    e = (row[4] if len(row) > 4 else "").strip()
    f = (row[5] if len(row) > 5 else "").strip()
    if _SKU_MASTER_DATE_RE.match(d) and _SKU_MASTER_DATE_RE.match(e) and _SKU_MASTER_PO_RE.match(f):
        return "old"
    if not d and not e and not f:
        return "unknown"
    return "unknown"


def _try_auto_migrate_sku_master(ws) -> bool:
    """
    尝试自动修复 UK_SKU_Master 的表头/数据错位问题。

    Returns:
        True  - 已经成功修复（或者扫描后发现其实不需要修复），调用方可以放心继续。
        False - 遇到了规则识别不了的行，没有安全地自动修复，调用方应该转去
                走"报错并停止"的逻辑，提示人工用 migrate_uk_sku_master.py 处理。
    """
    all_values = _call_with_retry(ws.get_all_values)
    if not all_values:
        return False
    data_rows = all_values[1:]

    old_rows = []
    for idx, row in enumerate(data_rows, start=2):
        padded = row + [""] * max(0, 8 - len(row))
        kind = _classify_legacy_sku_master_row(padded)
        if kind == "unknown":
            return False
        if kind == "old":
            old_rows.append((idx, padded))

    batch_data = []
    for idx, row in old_rows:
        d, e, f = row[3], row[4], row[5]
        # D(税率)清空、E<-原D(首次入库日期)、F<-原E(最后更新日期)、G<-原F(来源PO)
        batch_data.append({"range": f"D{idx}:G{idx}", "values": [["", d, e, f]]})
    # 表头统一成跟 SKU_MASTER_HEADERS 一致的 7 列，并清空多余的第 8 列
    # （旧表头里被误放在第 8 列的"税率"文字）
    batch_data.append({"range": "A1:H1", "values": [SKU_MASTER_HEADERS + [""]]})

    _call_with_retry(ws.batch_update, batch_data, value_input_option="USER_ENTERED")

    if old_rows:
        st.info(
            f"ℹ️ 检测到 UK_SKU_Master 的表头与代码字段顺序不一致（历史遗留问题：'税率'"
            f"功能上线时表头没有同步更新过），已自动修复：迁移了 {len(old_rows)} 行旧格式"
            f"数据、并更正了表头顺序。这条提示只会在修复当次出现，之后不会再提示。"
        )
    return True


def _fail_on_header_mismatch(actual_headers: list, expected_headers: list, sheet_title: str):
    """
    表头跟代码定义的顺序对不上、又没能自动修复时，与其让程序继续往错误的列
    写数据，不如在这里就停下来，把问题摆在明处。
    """
    st.error(
        f"⚠️ 「{sheet_title}」这张表的表头（第 1 行）跟代码里定义的字段顺序不一致，"
        f"为了避免继续往错误的列写入数据，程序已经停止运行。\n\n"
        f"代码期望的顺序：{expected_headers}\n\n"
        f"表里实际的顺序：{actual_headers}\n\n"
        f"请参考 migrate_uk_sku_master.py 里的说明手动排查、迁移数据，"
        f"或者手动把 Google Sheet 里这一行改回代码期望的顺序。"
    )
    st.stop()


def _get_or_create_worksheet(spreadsheet, title: str, headers: list):
    try:
        ws = _call_with_retry(spreadsheet.worksheet, title)
    except gspread.exceptions.WorksheetNotFound:
        ws = _call_with_retry(spreadsheet.add_worksheet, title=title, rows=1000, cols=len(headers) + 2)
        _call_with_retry(ws.append_row, headers)
        return ws

    existing_values = _call_with_retry(ws.row_values, 1)
    if not existing_values:
        _call_with_retry(ws.append_row, headers)
        return ws

    if existing_values[: len(headers)] != headers:
        if title == SKU_MASTER_SHEET_NAME and _try_auto_migrate_sku_master(ws):
            return ws
        _fail_on_header_mismatch(existing_values, headers, title)
    return ws


@st.cache_resource
def _get_sku_master_worksheet():
    """
    缓存住 UK_SKU_Master 这个 worksheet 的句柄。

    背景：_get_or_create_worksheet 每次被调用都要发 2 个读请求（按名字查找
    worksheet + 检查表头行是否存在），而这套代码里几乎每个数据库操作函数
    （读主数据库、写新 SKU、改记录、批量补税率……）都各自独立调用了一次——
    也就是说光是"找到这张表"这个动作，一次交互里就可能被重复发起好几次
    请求，很容易在使用高峰期打满 Google 的「每分钟读请求数」配额（这正是
    实测中真实报错的原因：APIError 429, Read requests per minute per user）。
    worksheet 的身份基本不会变，缓存成 cache_resource（跟应用进程同生命周期），
    这个查找动作整个部署周期只会真正发生一次。
    """
    return _get_or_create_worksheet(_get_spreadsheet(), SKU_MASTER_SHEET_NAME, SKU_MASTER_HEADERS)


def format_sku_master_sheet() -> dict:
    """
    给 UK_SKU_Master 做一次显示效果优化：冻结表头行、把列宽调整到能完整显示
    每列的内容、给有数据的区域整体加上边框。纯视觉/排版调整，不改变任何
    数据本身，可以随时重复调用（幂等，不会因为重复执行而堆积/出错）。

    设计成一个可以被 App 里的按钮直接调用的公开函数（而不是像之前那版一样
    藏在 _get_sku_master_worksheet 里"只在 worksheet 句柄第一次被创建时自动
    跑一次"）——那种写法的问题是：它到底有没有真的执行成功，用户在页面上
    完全看不出来（st.cache_resource 是跨 session 共享的，那次唯一的执行可能
    发生在任何一个用户的请求里，触发它的人也不一定会看到过程中的提示，而且
    st.cache_resource 命中缓存后就再也不会重新执行）。改成手动按钮触发之后，
    每次点击都能立刻看到明确的成功/失败反馈，出问题也能马上定位是哪一步、
    什么错误，而不是"看起来没生效但不知道为什么"。

    Returns:
        成功：{"success": True, "rows_formatted": N, "range": "A1:G123" 或 None（空表）}
        失败：{"success": False, "error": "具体的错误信息"}
        不会向外抛异常——调用方（按钮点击）直接把返回结果展示给用户即可。
    """
    try:
        ws = _get_sku_master_worksheet()
        all_values = _call_with_retry(ws.get_all_values)
        data_row_count = max(0, len(all_values) - 1)
        if data_row_count == 0:
            return {"success": True, "rows_formatted": 0, "range": None}

        last_row = data_row_count + 1  # +1 是表头行
        last_col_letter = gspread.utils.rowcol_to_a1(1, len(SKU_MASTER_HEADERS)).rstrip("0123456789")
        data_range = f"A1:{last_col_letter}{last_row}"

        # 冻结表头行，往下滚动的时候表头始终可见
        _call_with_retry(ws.freeze, rows=1)
        # 把每一列的宽度调整到刚好能完整显示该列最长的内容（Google Sheets 的
        # "调整到符合数据大小"），避免内容被相邻列挡住只显示一半
        _call_with_retry(ws.columns_auto_resize, 0, len(SKU_MASTER_HEADERS))
        # 给整个有数据的区域（表头 + 所有数据行）加上四周边框；显式指定
        # width/color，避免只写 style 在某些情况下渲染不出可见边框线
        border_side = {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}
        _call_with_retry(
            ws.format,
            data_range,
            {
                "borders": {
                    "top": border_side, "bottom": border_side,
                    "left": border_side, "right": border_side,
                }
            },
        )
        return {"success": True, "rows_formatted": data_row_count, "range": data_range}
    except Exception as e:
        return {"success": False, "error": f"{type(e).__name__}: {e}"}


@st.cache_resource
def _get_change_log_worksheet():
    """缓存住 UK_Change_Log 这个 worksheet 的句柄，理由同 _get_sku_master_worksheet。"""
    return _get_or_create_worksheet(_get_spreadsheet(), CHANGE_LOG_SHEET_NAME, CHANGE_LOG_HEADERS)


@st.cache_data(ttl=30, show_spinner=False)
def load_sku_master() -> pd.DataFrame:
    """
    读取 UK 主数据库，返回 DataFrame。若表为空，返回带正确列名的空 DataFrame。
    缓存 30 秒——Streamlit 每次交互（哪怕点的是无关按钮）都会重新执行整个脚本，
    没有缓存的话每次都要重新读一遍整张表，很容易触发 Google Sheets 的 429 读配额限制。
    写入操作（新增/更新 SKU）之后会主动清掉这个缓存，保证下次读到的是最新数据。
    """
    ws = _get_sku_master_worksheet()
    records = _call_with_retry(ws.get_all_records)
    if not records:
        return pd.DataFrame(columns=SKU_MASTER_HEADERS)

    df = pd.DataFrame(records)
    for col in SKU_MASTER_HEADERS:
        if col not in df.columns:
            df[col] = ""
    df["sku"] = df["sku"].astype(str).str.strip()
    df["hts"] = df["hts"].astype(str).str.strip()
    return df


@st.cache_data(ttl=30, show_spinner=False)
def load_change_log() -> pd.DataFrame:
    """同样缓存 30 秒，理由跟 load_sku_master 一样：避免每次脚本重跑都重新读表触发 429。"""
    ws = _get_change_log_worksheet()
    records = _call_with_retry(ws.get_all_records)
    if not records:
        return pd.DataFrame(columns=CHANGE_LOG_HEADERS)

    df = pd.DataFrame(records)
    # old_value / new_value 这两列可能混着数字（比如 HTS 编码）和文字（比如
    # 税率百分比字符串），Google Sheets 按"看起来像不像数字"逐个单元格自动
    # 判断类型，同一列里就可能一部分是 int、一部分是 str。这种混合类型的
    # object 列，Streamlit 用 st.dataframe 展示时转成 Arrow 表会报错
    # （ArrowTypeError: Expected bytes, got a 'int' object）。统一转成字符串，
    # 展示不受影响，也不会再报错。
    for col in ("old_value", "new_value"):
        if col in df.columns:
            df[col] = df[col].apply(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
    return df


def append_new_skus(new_skus: pd.DataFrame, source_po: str):
    """将全新 SKU（UK 数据库里原本不存在的）批量写入 UK_SKU_Master。"""
    if new_skus.empty:
        return

    ws = _get_sku_master_worksheet()

    today = datetime.now().strftime("%Y-%m-%d")
    rows = []
    for _, row in new_skus.iterrows():
        tax_rate = row["tax_rate"] if "tax_rate" in row and pd.notna(row["tax_rate"]) else ""
        rows.append([
            row["sku"], row["description"], row["hts"], tax_rate,
            today, today, source_po or "",
        ])

    _call_with_retry(ws.append_rows, rows, value_input_option="USER_ENTERED")
    load_sku_master.clear()


def update_sku_record(sku: str, new_description: str, new_hts: str, source_po: str,
                       new_tax_rate: str | None = None):
    """人工确认后，用新值覆盖 UK 主数据库中该 SKU 的记录。
    new_tax_rate 为 None 或空字符串时不覆盖数据库里已有的税率（本次没读到税率不代表税率变了）。

    这个函数每次调用要发好几个 Google Sheets API 请求（1 次读表 + 最多 4 次
    单元格写入），单条人工确认点一次按钮没问题，但不适合在循环里对着一批
    SKU 连续调用——很容易在几秒内打满 Google 的每分钟请求配额，触发 429
    报错（APIError）。批量场景请用 bulk_update_tax_rates。
    """
    ws = _get_sku_master_worksheet()

    all_values = _call_with_retry(ws.get_all_values)
    headers = all_values[0]
    sku_col_idx = headers.index("sku")
    desc_col_idx = headers.index("description")
    hts_col_idx = headers.index("hts")
    last_updated_idx = headers.index("last_updated_date")
    tax_col_idx = headers.index("tax_rate") if "tax_rate" in headers else None

    today = datetime.now().strftime("%Y-%m-%d")

    for row_idx, row in enumerate(all_values[1:], start=2):
        if len(row) > sku_col_idx and row[sku_col_idx].strip() == sku.strip():
            # 这几个单元格改动合并成一次 batch_update，而不是 4 次单独的
            # update_cell 调用——减少请求数，也降低撞上配额限制的概率。
            cell_updates = [
                {"range": gspread.utils.rowcol_to_a1(row_idx, desc_col_idx + 1), "values": [[new_description]]},
                {"range": gspread.utils.rowcol_to_a1(row_idx, hts_col_idx + 1), "values": [[new_hts]]},
                {"range": gspread.utils.rowcol_to_a1(row_idx, last_updated_idx + 1), "values": [[today]]},
            ]
            if new_tax_rate and tax_col_idx is not None:
                cell_updates.append(
                    {"range": gspread.utils.rowcol_to_a1(row_idx, tax_col_idx + 1), "values": [[new_tax_rate]]}
                )
            _call_with_retry(ws.batch_update, cell_updates, value_input_option="USER_ENTERED")
            load_sku_master.clear()
            return True
    return False


def bulk_update_tax_rates(updates: list[dict]) -> int:
    """
    批量把税率写入 UK_SKU_Master，用于"税率从空值自动补全"这种一次上传里
    可能牵涉几十个 SKU 的场景。

    跟 update_sku_record 逐个 SKU 调用不同，这里只发 3 个 Google Sheets API
    请求（打开表 + 读一次全表 + 一次性批量写入所有改动的单元格），不管
    updates 里有多少条，请求数都不会涨——避免像逐条调用 update_sku_record
    那样，在一批几十个 SKU 的场景里几秒内打满 Google 的每分钟请求配额、
    触发 429 报错。

    Args:
        updates: [{"sku": ..., "tax_rate": "2.7%"}, ...]

    Returns:
        实际找到并更新的行数（在表里找不到的 SKU 会被跳过，不报错）。
    """
    if not updates:
        return 0

    ws = _get_sku_master_worksheet()

    all_values = _call_with_retry(ws.get_all_values)
    if not all_values:
        return 0
    headers = all_values[0]
    if "sku" not in headers or "tax_rate" not in headers:
        return 0
    sku_col_idx = headers.index("sku")
    tax_col_idx = headers.index("tax_rate")
    last_updated_idx = headers.index("last_updated_date") if "last_updated_date" in headers else None

    sku_to_row = {}
    for row_idx, row in enumerate(all_values[1:], start=2):
        if len(row) > sku_col_idx and row[sku_col_idx].strip():
            sku_to_row[row[sku_col_idx].strip()] = row_idx

    today = datetime.now().strftime("%Y-%m-%d")
    batch_data = []
    updated_count = 0
    for u in updates:
        row_idx = sku_to_row.get(str(u["sku"]).strip())
        if row_idx is None:
            continue
        batch_data.append({
            "range": gspread.utils.rowcol_to_a1(row_idx, tax_col_idx + 1),
            "values": [[u["tax_rate"]]],
        })
        if last_updated_idx is not None:
            batch_data.append({
                "range": gspread.utils.rowcol_to_a1(row_idx, last_updated_idx + 1),
                "values": [[today]],
            })
        updated_count += 1

    if batch_data:
        _call_with_retry(ws.batch_update, batch_data, value_input_option="USER_ENTERED")
        load_sku_master.clear()

    return updated_count


def append_change_log(entries: list[dict]):
    """批量写入 UK 变更日志。"""
    if not entries:
        return

    ws = _get_change_log_worksheet()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for e in entries:
        rows.append([
            timestamp, e["sku"], e["field_changed"], e["old_value"],
            e["new_value"], e.get("source_po", ""), e.get("resolution", "pending"),
        ])

    _call_with_retry(ws.append_rows, rows, value_input_option="USER_ENTERED")
    load_change_log.clear()
