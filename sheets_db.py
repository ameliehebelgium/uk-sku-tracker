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
    return client.open_by_key(sheet_id)


def _get_or_create_worksheet(spreadsheet, title: str, headers: list):
    try:
        ws = spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=1000, cols=len(headers) + 2)
        ws.append_row(headers)
        return ws

    existing_values = ws.row_values(1)
    if not existing_values:
        ws.append_row(headers)
    return ws


@st.cache_data(ttl=30, show_spinner=False)
def load_sku_master() -> pd.DataFrame:
    """
    读取 UK 主数据库，返回 DataFrame。若表为空，返回带正确列名的空 DataFrame。
    缓存 30 秒——Streamlit 每次交互（哪怕点的是无关按钮）都会重新执行整个脚本，
    没有缓存的话每次都要重新读一遍整张表，很容易触发 Google Sheets 的 429 读配额限制。
    写入操作（新增/更新 SKU）之后会主动清掉这个缓存，保证下次读到的是最新数据。
    """
    spreadsheet = _get_spreadsheet()
    ws = _get_or_create_worksheet(spreadsheet, SKU_MASTER_SHEET_NAME, SKU_MASTER_HEADERS)
    records = ws.get_all_records()
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
    spreadsheet = _get_spreadsheet()
    ws = _get_or_create_worksheet(spreadsheet, CHANGE_LOG_SHEET_NAME, CHANGE_LOG_HEADERS)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=CHANGE_LOG_HEADERS)
    return pd.DataFrame(records)


def append_new_skus(new_skus: pd.DataFrame, source_po: str):
    """将全新 SKU（UK 数据库里原本不存在的）批量写入 UK_SKU_Master。"""
    if new_skus.empty:
        return

    spreadsheet = _get_spreadsheet()
    ws = _get_or_create_worksheet(spreadsheet, SKU_MASTER_SHEET_NAME, SKU_MASTER_HEADERS)

    today = datetime.now().strftime("%Y-%m-%d")
    rows = []
    for _, row in new_skus.iterrows():
        tax_rate = row["tax_rate"] if "tax_rate" in row and pd.notna(row["tax_rate"]) else ""
        rows.append([
            row["sku"], row["description"], row["hts"], tax_rate,
            today, today, source_po or "",
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    load_sku_master.clear()


def update_sku_record(sku: str, new_description: str, new_hts: str, source_po: str,
                       new_tax_rate: str | None = None):
    """人工确认后，用新值覆盖 UK 主数据库中该 SKU 的记录。
    new_tax_rate 为 None 或空字符串时不覆盖数据库里已有的税率（本次没读到税率不代表税率变了）。
    """
    spreadsheet = _get_spreadsheet()
    ws = _get_or_create_worksheet(spreadsheet, SKU_MASTER_SHEET_NAME, SKU_MASTER_HEADERS)

    all_values = ws.get_all_values()
    headers = all_values[0]
    sku_col_idx = headers.index("sku")
    desc_col_idx = headers.index("description")
    hts_col_idx = headers.index("hts")
    last_updated_idx = headers.index("last_updated_date")
    tax_col_idx = headers.index("tax_rate") if "tax_rate" in headers else None

    today = datetime.now().strftime("%Y-%m-%d")

    for row_idx, row in enumerate(all_values[1:], start=2):
        if len(row) > sku_col_idx and row[sku_col_idx].strip() == sku.strip():
            ws.update_cell(row_idx, desc_col_idx + 1, new_description)
            ws.update_cell(row_idx, hts_col_idx + 1, new_hts)
            if new_tax_rate and tax_col_idx is not None:
                ws.update_cell(row_idx, tax_col_idx + 1, new_tax_rate)
            ws.update_cell(row_idx, last_updated_idx + 1, today)
            load_sku_master.clear()
            return True
    return False


def append_change_log(entries: list[dict]):
    """批量写入 UK 变更日志。"""
    if not entries:
        return

    spreadsheet = _get_spreadsheet()
    ws = _get_or_create_worksheet(spreadsheet, CHANGE_LOG_SHEET_NAME, CHANGE_LOG_HEADERS)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for e in entries:
        rows.append([
            timestamp, e["sku"], e["field_changed"], e["old_value"],
            e["new_value"], e.get("source_po", ""), e.get("resolution", "pending"),
        ])

    ws.append_rows(rows, value_input_option="USER_ENTERED")
    load_change_log.clear()
