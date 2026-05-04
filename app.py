"""
NailVesta 中台运营周报生成器 v2 ✨
==================================
作者：为 Chenhao（奶瓶）打造

输入（每周固定上传 3 个文件）：
    1. 28 天订单数据（All_order csv/xlsx）
    2. 28 天退货数据（Return_Refund_Orders xlsx）
    3. 产品图册（NailVesta_产品图册 csv） — 用于识别供应商和新品（≤30 天）

时间窗模式（侧边栏可切换）：
    A. 28 天对半切：前 14 天 vs 后 14 天（默认，数据完整）
    B. 自然周对比：最近完整一周（周一至周日）vs 上一周

启动：
    streamlit run app.py
"""

import io
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 1 部分：常量配置                                              ║
# ╚════════════════════════════════════════════════════════════════════╝

INFLUENCER_RULE_MAP = {
    "SKU Subtotal After Discount = 0": "SKU Subtotal After Discount",
    "SKU Unit Original Price = 0": "SKU Unit Original Price",
    "Order Amount = 0": "Order Amount",
}

# 模板里要复制的基础 sheet（最新版字段最完整）
BASE_SHEET = "4.20-4.26"

# 新品阈值（天数）
NEW_PRODUCT_DAYS = 30

# 退货文件列名候选（兼容多版本）
RETURN_COL_CANDIDATES = {
    "qty":              ["Return Quantity", "Sku Quantity of return", "Quantity", "退货数量"],
    "unit_price":       ["Return unit price", "Unit Price", "退款单价"],
    "amount":           ["Order Refund Amount", "Refund Amount", "退款金额"],
    "reason":           ["Return Reason", "Refund Reason", "Reason", "退货原因"],
    "sku_name":         ["SKU Name", "Variation", "SKU Variation", "款式"],
    "sku":              ["Seller SKU", "SKU", "Seller_SKU"],
    "return_status":    ["Return Status", "退货状态"],
    "return_substatus": ["Return Sub Status", "Return SubStatus", "退货子状态"],
    "time_requested":   ["Time Requested", "申请时间"],
    "order_id":         ["Order ID", "OrderID", "订单号"],
}

# 「不计入退货指标」的子状态：买家主动撤销，跟客服无关
EXCLUDED_SUBSTATUS_KEYWORDS = ["request canceled", "request cancelled"]

# 产品图册列名（兼容编辑差异）
CATALOG_COL_CANDIDATES = {
    "sku":          ["SKU", "sku"],
    "style_en":     ["款式英文名称", "英文名称", "Style"],
    "supplier":     ["厂家", "供应商", "Supplier"],
    "list_date":    ["上架时间", "上架日期", "Listed Date"],
    "list_status":  ["上架状态", "状态"],
    "is_listed":    ["是否上架"],
}


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 2 部分：通用工具函数                                          ║
# ╚════════════════════════════════════════════════════════════════════╝

def _read_any(file) -> pd.DataFrame:
    """读 csv 或 xlsx，自动识别。"""
    if file is None:
        return pd.DataFrame()
    name = file.name.lower()
    file.seek(0)
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


def _find_col(df: pd.DataFrame, candidates: list) -> str:
    """从候选列名里找到第一个存在的，返回 None 表示都不存在。"""
    for c in candidates:
        if c in df.columns:
            return c
    return None


def extract_style(variation) -> str:
    """从 Variation/SKU Name 提取款式名（"Acai Bloom, L" → "Acai Bloom"）。"""
    if pd.isna(variation):
        return ""
    return str(variation).split(",")[0].strip()


def extract_sku_root(sku) -> str:
    """从 Seller SKU 提取根 SKU（"NOF027-L" → "NOF027"）。"""
    if pd.isna(sku):
        return ""
    return str(sku).split("-")[0].strip()


def _bucket_sku(n: int) -> str:
    if n == 1: return "1"
    if n == 2: return "2"
    if n == 3: return "3"
    if n == 4: return "4"
    return "4+"


def _wow(this: float, last: float) -> float:
    if not last:
        return 0
    return this / last - 1


def _parse_pct(s: str):
    if not s:
        return None
    try:
        v = float(s.replace("%", "").strip())
        return v / 100
    except ValueError:
        return None


def _is_request_canceled(substatus_value) -> bool:
    if pd.isna(substatus_value):
        return False
    s = str(substatus_value).lower()
    return any(kw in s for kw in EXCLUDED_SUBSTATUS_KEYWORDS)


def _clean_tab_str(s: pd.Series) -> pd.Series:
    """订单/退货 csv 时间字段经常带 \\t，统一清理。"""
    return s.astype(str).str.strip().str.rstrip("\t").str.strip()


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 3 部分：时间窗拆分                                            ║
# ╚════════════════════════════════════════════════════════════════════╝

def split_time_windows(orders_df: pd.DataFrame, returns_df: pd.DataFrame,
                       mode: str) -> dict:
    """
    将 28 天数据拆为「上周/本周」两段。

    mode:
        "halve"      — 28 天对半切（前 14 = 上周，后 14 = 本周）
        "weekly"     — 自然周对比（最近完整周一-周日 vs 上一周）

    返回：
        {
          "last_orders": df, "this_orders": df,
          "last_returns": df, "this_returns": df,
          "last_label": str, "this_label": str,
          "ref_date": datetime  # 用于新品识别的基准日
        }
    """
    if orders_df.empty:
        raise ValueError("订单数据为空，无法切分")

    # 解析订单时间（Created Time 通常带 \t）
    o = orders_df.copy()
    if "Created Time" in o.columns:
        o["__t"] = pd.to_datetime(_clean_tab_str(o["Created Time"]), errors="coerce")
    else:
        raise ValueError("订单文件缺少 Created Time 列")

    # 解析退货时间（Time Requested 是 DD/MM/YYYY 格式）
    r = returns_df.copy() if not returns_df.empty else pd.DataFrame()
    if not r.empty:
        time_col = _find_col(r, RETURN_COL_CANDIDATES["time_requested"])
        if time_col:
            cleaned = _clean_tab_str(r[time_col])
            # 优先按 DD/MM/YYYY，失败再回退到自动识别
            r["__t"] = pd.to_datetime(cleaned, format="%d/%m/%Y %H:%M:%S", errors="coerce")
            mask_na = r["__t"].isna()
            if mask_na.any():
                r.loc[mask_na, "__t"] = pd.to_datetime(cleaned[mask_na], errors="coerce")
        else:
            r["__t"] = pd.NaT

    # 时间范围
    t_min = o["__t"].min()
    t_max = o["__t"].max()
    if pd.isna(t_min) or pd.isna(t_max):
        raise ValueError("订单时间字段无法解析")

    if mode == "halve":
        # 28 天对半切
        midpoint = t_min + (t_max - t_min) / 2
        last_start, last_end = t_min, midpoint
        this_start, this_end = midpoint, t_max + timedelta(seconds=1)
        last_label  = f"{last_start.strftime('%-m.%-d')}-{(last_end - timedelta(days=1)).strftime('%-m.%-d')}"
        this_label  = f"{this_start.strftime('%-m.%-d')}-{t_max.strftime('%-m.%-d')}"

    elif mode == "weekly":
        # 找最近一个周日作为本周末（如果 t_max 已经是周日，就用它）
        # weekday(): 周一=0, 周日=6
        days_to_sunday = (6 - t_max.weekday()) % 7
        # 如果 t_max 不是周日，往前找最近的周日（结束于上周日）
        if t_max.weekday() != 6:
            this_end = t_max - timedelta(days=t_max.weekday() + 1)
            this_end = this_end.replace(hour=23, minute=59, second=59)
        else:
            this_end = t_max
        this_end = pd.Timestamp(this_end.replace(hour=23, minute=59, second=59))
        this_start = (this_end - timedelta(days=6)).replace(hour=0, minute=0, second=0)
        last_end = (this_start - timedelta(seconds=1))
        last_start = (last_end - timedelta(days=6)).replace(hour=0, minute=0, second=0)
        last_label = f"{last_start.strftime('%-m.%-d')}-{last_end.strftime('%-m.%-d')}"
        this_label = f"{this_start.strftime('%-m.%-d')}-{this_end.strftime('%-m.%-d')}"
    else:
        raise ValueError(f"未知时间窗模式: {mode}")

    # 切订单
    last_orders = o[(o["__t"] >= last_start) & (o["__t"] <= last_end)].copy()
    this_orders = o[(o["__t"] >= this_start) & (o["__t"] <= this_end)].copy()

    # 切退货
    if not r.empty and "__t" in r.columns:
        last_returns = r[(r["__t"] >= last_start) & (r["__t"] <= last_end)].copy()
        this_returns = r[(r["__t"] >= this_start) & (r["__t"] <= this_end)].copy()
    else:
        last_returns = pd.DataFrame()
        this_returns = pd.DataFrame()

    return {
        "last_orders": last_orders,
        "this_orders": this_orders,
        "last_returns": last_returns,
        "this_returns": this_returns,
        "last_label": last_label,
        "this_label": this_label,
        "ref_date": t_max,  # 用于新品识别
        "t_min": t_min, "t_max": t_max,
    }


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 4 部分：产品图册处理                                          ║
# ╚════════════════════════════════════════════════════════════════════╝

def process_catalog(df: pd.DataFrame) -> pd.DataFrame:
    """
    清洗产品图册，返回标准格式：
        sku_root | style_en | supplier | list_date | is_active
    """
    if df.empty:
        return pd.DataFrame(columns=["sku_root", "style_en", "supplier", "list_date", "is_active"])

    sku_col      = _find_col(df, CATALOG_COL_CANDIDATES["sku"])
    style_col    = _find_col(df, CATALOG_COL_CANDIDATES["style_en"])
    supplier_col = _find_col(df, CATALOG_COL_CANDIDATES["supplier"])
    list_col     = _find_col(df, CATALOG_COL_CANDIDATES["list_date"])
    status_col   = _find_col(df, CATALOG_COL_CANDIDATES["list_status"])
    is_listed_col= _find_col(df, CATALOG_COL_CANDIDATES["is_listed"])

    out = pd.DataFrame()
    out["sku_root"]  = df[sku_col].astype(str).str.strip() if sku_col else ""
    out["style_en"]  = df[style_col].astype(str).str.strip() if style_col else ""
    out["supplier"]  = df[supplier_col].astype(str).str.strip() if supplier_col else ""
    out["list_date"] = pd.to_datetime(df[list_col], errors="coerce") if list_col else pd.NaT
    if status_col:
        out["status"] = df[status_col].astype(str).str.strip()
    else:
        out["status"] = ""
    if is_listed_col:
        out["is_active"] = pd.to_numeric(df[is_listed_col], errors="coerce").fillna(0) > 0
    else:
        out["is_active"] = True

    # 清空 supplier 里的 nan/空
    out["supplier"] = out["supplier"].replace({"nan": "", "NaN": "", "None": ""})

    return out


def is_new_product(catalog: pd.DataFrame, ref_date: datetime,
                   days: int = NEW_PRODUCT_DAYS) -> pd.DataFrame:
    """筛选 ≤days 天的新品，返回带新品标记的图册。"""
    if catalog.empty:
        return catalog
    threshold = ref_date - timedelta(days=days)
    catalog = catalog.copy()
    catalog["is_new"] = catalog["list_date"] >= threshold
    return catalog


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 5 部分：订单数据处理                                          ║
# ╚════════════════════════════════════════════════════════════════════╝

def _empty_orders() -> dict:
    return {
        "orders": 0, "sku_sold": 0, "gmv": 0,
        "aov": 0, "asp": 0, "attach_rate": 0,
        "influencer_orders": 0,
        "bucket_struct": pd.DataFrame(columns=["bucket", "order_count"]),
        "bucket_aov": {},
        "style_sales": pd.DataFrame(columns=["style", "sales", "sku"]),
        "raw_paid": pd.DataFrame(),
    }


def process_orders(df: pd.DataFrame, influencer_col: str) -> dict:
    """剔除达人单后，输出核心指标。"""
    if df.empty:
        return _empty_orders()

    df = df.copy()

    for col in ["SKU Subtotal After Discount", "SKU Unit Original Price",
                "Order Amount", "Quantity"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["is_influencer"] = df[influencer_col] == 0

    if "Order Status" in df.columns:
        df_valid = df[df["Order Status"] != "Canceled"].copy()
    else:
        df_valid = df.copy()

    paid = df_valid[~df_valid["is_influencer"]].copy()
    influencer = df_valid[df_valid["is_influencer"]].copy()

    paid["style"] = paid["Variation"].apply(extract_style)
    paid["sku_root"] = paid["Seller SKU"].apply(extract_sku_root)
    influencer["style"] = influencer["Variation"].apply(extract_style)

    paid_orders = paid["Order ID"].nunique()
    paid_sku_sold = int(paid["Quantity"].sum())
    paid_gmv = float(paid["SKU Subtotal After Discount"].sum())

    aov = paid_gmv / paid_orders if paid_orders else 0
    asp = paid_gmv / paid_sku_sold if paid_sku_sold else 0
    attach_rate = paid_sku_sold / paid_orders if paid_orders else 0

    sku_per_order = paid.groupby("Order ID")["Quantity"].sum().rename("sku_count").reset_index()
    sku_per_order["bucket"] = sku_per_order["sku_count"].apply(_bucket_sku)
    bucket_struct = sku_per_order.groupby("bucket").agg(order_count=("Order ID", "count")).reset_index()

    order_amt = paid.groupby("Order ID")["SKU Subtotal After Discount"].sum().rename("order_amt").reset_index()
    sku_per_order = sku_per_order.merge(order_amt, on="Order ID", how="left")
    bucket_aov = sku_per_order.groupby("bucket")["order_amt"].mean().to_dict()

    # 销量 TOP（含 sku_root，方便后面关联图册）
    style_sales = (
        paid.groupby("style")
        .agg(
            sales=("Quantity", "sum"),
            sku=("sku_root", lambda s: s.mode().iloc[0] if not s.mode().empty else ""),
        )
        .reset_index()
        .sort_values("sales", ascending=False)
    )

    return {
        "orders": paid_orders,
        "sku_sold": paid_sku_sold,
        "gmv": paid_gmv,
        "aov": aov, "asp": asp,
        "attach_rate": attach_rate,
        "influencer_orders": influencer["Order ID"].nunique(),
        "bucket_struct": bucket_struct,
        "bucket_aov": bucket_aov,
        "style_sales": style_sales,
        "raw_paid": paid,
    }


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 6 部分：退货数据处理                                          ║
# ╚════════════════════════════════════════════════════════════════════╝

def _empty_returns() -> dict:
    return {
        "total_qty": 0, "total_amount": 0,
        "by_style": pd.DataFrame(columns=["style", "return_qty"]),
        "by_reason": pd.DataFrame(columns=["reason", "qty"]),
        "by_reason_top_style": {},
        "requested_qty": 0,
        "request_canceled_qty": 0,
    }


def process_returns(df: pd.DataFrame, exclude_request_canceled: bool = True) -> dict:
    """
    退货处理双层逻辑：
    - 实际退货量 / 金额 / by_style → 排除 Request Canceled（买家撤销）
      但保留 Refund rejected（客服拒绝）
    - 退货原因 → 含全部（含撤销和被拒）
    """
    if df.empty:
        return _empty_returns()

    df = df.copy()

    qty_col       = _find_col(df, RETURN_COL_CANDIDATES["qty"])
    unit_price_col= _find_col(df, RETURN_COL_CANDIDATES["unit_price"])
    amt_col       = _find_col(df, RETURN_COL_CANDIDATES["amount"])
    reason_col    = _find_col(df, RETURN_COL_CANDIDATES["reason"])
    sku_name_col  = _find_col(df, RETURN_COL_CANDIDATES["sku_name"])
    sku_col       = _find_col(df, RETURN_COL_CANDIDATES["sku"])
    substatus_col = _find_col(df, RETURN_COL_CANDIDATES["return_substatus"])

    if qty_col:
        df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
        df = df[df[qty_col] > 0].copy()
    else:
        df["__qty"] = 1
        qty_col = "__qty"

    if sku_name_col:
        df["style"] = df[sku_name_col].apply(extract_style)
    elif sku_col:
        df["style"] = df[sku_col].apply(extract_sku_root)
    else:
        df["style"] = ""

    requested_qty = int(df[qty_col].sum())

    if substatus_col:
        df["__is_req_canceled"] = df[substatus_col].apply(_is_request_canceled)
    else:
        df["__is_req_canceled"] = False

    request_canceled_qty = int(df[df["__is_req_canceled"]][qty_col].sum())

    if exclude_request_canceled:
        df_actual = df[~df["__is_req_canceled"]].copy()
    else:
        df_actual = df.copy()

    total_qty = int(df_actual[qty_col].sum()) if not df_actual.empty else 0

    if not df_actual.empty:
        if unit_price_col:
            df_actual[unit_price_col] = pd.to_numeric(df_actual[unit_price_col], errors="coerce").fillna(0)
            total_amount = float((df_actual[unit_price_col] * df_actual[qty_col]).sum())
        elif amt_col:
            df_actual[amt_col] = pd.to_numeric(df_actual[amt_col], errors="coerce").fillna(0)
            total_amount = float(df_actual[amt_col].sum())
        else:
            total_amount = 0
    else:
        total_amount = 0

    if not df_actual.empty:
        by_style = (
            df_actual.groupby("style")[qty_col].sum().rename("return_qty").reset_index()
            .sort_values("return_qty", ascending=False)
        )
    else:
        by_style = pd.DataFrame(columns=["style", "return_qty"])

    if reason_col:
        by_reason = (
            df.groupby(reason_col)[qty_col].sum().rename("qty").reset_index()
            .rename(columns={reason_col: "reason"})
            .sort_values("qty", ascending=False)
        )
        by_reason_top = {}
        for r, sub in df.groupby(reason_col):
            top = sub.groupby("style")[qty_col].sum().sort_values(ascending=False).head(2)
            by_reason_top[r] = "; ".join([f"{s} ({int(q)})" for s, q in top.items() if s])
    else:
        by_reason = pd.DataFrame(columns=["reason", "qty"])
        by_reason_top = {}

    return {
        "total_qty": total_qty,
        "total_amount": total_amount,
        "by_style": by_style,
        "by_reason": by_reason,
        "by_reason_top_style": by_reason_top,
        "requested_qty": requested_qty,
        "request_canceled_qty": request_canceled_qty,
    }


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 7 部分：新品 + 供应商分析（基于图册）                         ║
# ╚════════════════════════════════════════════════════════════════════╝

def build_new_products_top10(this_orders: dict, this_returns: dict,
                              catalog: pd.DataFrame, ref_date: datetime) -> pd.DataFrame:
    """
    近 30 天新品销量 TOP 10。
    输出列：style | sku | sales | return_qty | return_rate
    """
    if catalog.empty or "list_date" not in catalog.columns:
        return pd.DataFrame(columns=["style", "sku", "sales", "return_qty", "return_rate"])

    cat = is_new_product(catalog, ref_date)
    new_skus = cat[cat["is_new"]]["sku_root"].dropna().unique().tolist()
    if not new_skus:
        return pd.DataFrame(columns=["style", "sku", "sales", "return_qty", "return_rate"])

    paid = this_orders["raw_paid"]
    if paid.empty:
        return pd.DataFrame(columns=["style", "sku", "sales", "return_qty", "return_rate"])

    new_paid = paid[paid["sku_root"].isin(new_skus)].copy()
    if new_paid.empty:
        return pd.DataFrame(columns=["style", "sku", "sales", "return_qty", "return_rate"])

    style_sales = (
        new_paid.groupby(["style", "sku_root"])
        .agg(sales=("Quantity", "sum"))
        .reset_index()
        .rename(columns={"sku_root": "sku"})
    )
    style_sales = style_sales.merge(
        this_returns["by_style"], on="style", how="left"
    ).fillna({"return_qty": 0})
    style_sales["return_rate"] = style_sales.apply(
        lambda r: r["return_qty"] / r["sales"] if r["sales"] > 0 else 0, axis=1
    )
    return style_sales.sort_values("sales", ascending=False).head(10)[
        ["style", "sku", "sales", "return_qty", "return_rate"]
    ]


def build_supplier_problem(this_orders: dict, this_returns: dict,
                            catalog: pd.DataFrame,
                            return_rate_threshold: float = 0.10,
                            min_sales: int = 10) -> pd.DataFrame:
    """
    问题款式归因到供应商。
    把"退货率高 ≥ 10% 且销量 ≥ 10"的款定义为问题款，按供应商聚合。
    输出列：supplier | problem_count | ratio | top_style
    """
    if catalog.empty:
        return pd.DataFrame(columns=["supplier", "problem_count", "ratio", "top_style"])

    paid = this_orders["raw_paid"]
    if paid.empty:
        return pd.DataFrame(columns=["supplier", "problem_count", "ratio", "top_style"])

    # 款式销量 + 退货率
    style_sales = (
        paid.groupby(["style", "sku_root"])
        .agg(sales=("Quantity", "sum"))
        .reset_index()
    )
    style_sales = style_sales.merge(this_returns["by_style"], on="style", how="left").fillna({"return_qty": 0})
    style_sales["return_rate"] = style_sales.apply(
        lambda r: r["return_qty"] / r["sales"] if r["sales"] > 0 else 0, axis=1
    )

    # 关联图册（按 sku_root）
    cat_lookup = catalog[["sku_root", "supplier"]].drop_duplicates(subset=["sku_root"])
    style_sales = style_sales.merge(cat_lookup, on="sku_root", how="left")
    style_sales["supplier"] = style_sales["supplier"].fillna("未知").replace("", "未知")

    # 标记问题款
    is_problem = (style_sales["return_rate"] >= return_rate_threshold) & (style_sales["sales"] >= min_sales)
    problem_styles = style_sales[is_problem].copy()

    if problem_styles.empty:
        return pd.DataFrame(columns=["supplier", "problem_count", "ratio", "top_style"])

    total_problem = len(problem_styles)
    summary = (
        problem_styles.groupby("supplier")
        .agg(
            problem_count=("style", "count"),
            top_style=("style", lambda s: ", ".join(s.head(2).tolist())),
        )
        .reset_index()
        .sort_values("problem_count", ascending=False)
    )
    summary["ratio"] = summary["problem_count"] / total_problem
    return summary[["supplier", "problem_count", "ratio", "top_style"]]


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 8 部分：写入 Excel 模板                                       ║
# ╚════════════════════════════════════════════════════════════════════╝

def _write_template(template_bytes, last_label, this_label,
                    last_o, this_o, last_r, this_r,
                    top10_sales, top10_return,
                    new_products_top10, supplier_problem,
                    traffic_metrics) -> bytes:
    """把所有指标写到模板的副本里。"""
    wb = load_workbook(io.BytesIO(template_bytes))

    if BASE_SHEET in wb.sheetnames:
        base_ws = wb[BASE_SHEET]
    else:
        base_ws = wb[wb.sheetnames[-1]]

    ws = wb.copy_worksheet(base_ws)
    # sheet 名只能 ≤31 字符且不含特殊字符
    safe_name = this_label.replace(":", "_").replace("/", "-")[:31]
    ws.title = safe_name

    # ---------------- B2 标题 ----------------
    ws["B2"] = f"NailVesta 中台运营周报  |  Week：{this_label}"

    # ---------------- 核心指标 (B5:G15) ----------------
    metrics_rows = {
        6:  last_o["orders"],     7:  last_o["sku_sold"],
        8:  last_o["attach_rate"],9:  last_o["asp"],
        10: last_o["aov"],        15: last_o["gmv"],
    }
    this_metrics = {
        6:  this_o["orders"],     7:  this_o["sku_sold"],
        8:  this_o["attach_rate"],9:  this_o["asp"],
        10: this_o["aov"],        15: this_o["gmv"],
    }
    for row, last_v in metrics_rows.items():
        ws.cell(row=row, column=3, value=last_v)
        ws.cell(row=row, column=6, value=this_metrics[row])
        ws.cell(row=row, column=7, value=f"=IFERROR(F{row}/C{row}-1,\"\")")

    ws["D6"] = last_o["influencer_orders"]
    ws["E6"] = this_o["influencer_orders"]
    ws["D7"] = ""
    ws["E7"] = ""

    # 流量指标
    flow_map = {
        11: ("ctr_last", "ctr_this"),
        12: ("cvr_last", "cvr_this"),
        13: ("cart_conv_last", "cart_conv_this"),
        14: ("atc_last", "atc_this"),
    }
    for row, (lk, tk) in flow_map.items():
        lv = _parse_pct(traffic_metrics.get(lk, ""))
        tv = _parse_pct(traffic_metrics.get(tk, ""))
        if lv is not None:
            ws.cell(row=row, column=3, value=lv).number_format = "0.00%"
        if tv is not None:
            ws.cell(row=row, column=6, value=tv).number_format = "0.00%"

    # ---------------- 退换货总览 (L5:Q15) ----------------
    rr_last = last_r["total_qty"] / last_o["sku_sold"] if last_o["sku_sold"] else 0
    rr_this = this_r["total_qty"] / this_o["sku_sold"] if this_o["sku_sold"] else 0
    per_last = last_r["total_amount"] / last_r["total_qty"] if last_r["total_qty"] else 0
    per_this = this_r["total_amount"] / this_r["total_qty"] if this_r["total_qty"] else 0

    return_rows = {
        6:  (last_r["total_qty"],     this_r["total_qty"]),
        7:  (rr_last,                 rr_this),
        8:  (last_r["total_amount"],  this_r["total_amount"]),
        9:  (per_last,                per_this),
        10: (len(last_r["by_style"]), len(this_r["by_style"])),
    }
    for row, (lv, tv) in return_rows.items():
        ws.cell(row=row, column=13, value=lv)
        ws.cell(row=row, column=14, value=tv)
        ws.cell(row=row, column=15, value=f"=IFERROR(N{row}/M{row}-1,\"\")")
    ws["M7"].number_format = "0.00%"
    ws["N7"].number_format = "0.00%"

    # ---------------- 订单内 SKU 数量结构 ----------------
    bucket_rows = {"1": 18, "2": 19, "3": 20, "4": 21, "4+": 22}
    last_bucket = last_o["bucket_struct"].set_index("bucket")["order_count"].to_dict()
    this_bucket = this_o["bucket_struct"].set_index("bucket")["order_count"].to_dict()
    last_total = sum(last_bucket.values()) or 1
    this_total = sum(this_bucket.values()) or 1

    for b, row in bucket_rows.items():
        lo = last_bucket.get(b, 0)
        to = this_bucket.get(b, 0)
        ws.cell(row=row, column=3, value=lo)
        ws.cell(row=row, column=4, value=lo / last_total).number_format = "0.00%"
        ws.cell(row=row, column=5, value=round(last_o["bucket_aov"].get(b, 0), 2))
        ws.cell(row=row, column=6, value=to)
        ws.cell(row=row, column=7, value=to / this_total).number_format = "0.00%"
        ws.cell(row=row, column=8, value=round(this_o["bucket_aov"].get(b, 0), 2))

    # ---------------- 退货原因 (L17:Q26) ----------------
    reason_rows_map = {
        "No longer needed":                                       18,
        "Missing package":                                        19,
        "Wrong item was sent":                                    20,
        "Item doesn't match description":                         21,
        "Defective item":                                         22,
        "Product wouldn't arrive on time":                        23,
        "Congrats on meeting your refundable sample criteria!":   24,
        "Missing items":                                          25,
        "Damaged item or packaging":                              26,
    }
    last_reason = last_r["by_reason"].set_index("reason")["qty"].to_dict() if not last_r["by_reason"].empty else {}
    this_reason = this_r["by_reason"].set_index("reason")["qty"].to_dict() if not this_r["by_reason"].empty else {}
    last_reason_total = sum(last_reason.values()) or 1
    this_reason_total = sum(this_reason.values()) or 1

    def _match_reason(target: str, reason_dict: dict) -> int:
        target_l = target.lower()
        for k, v in reason_dict.items():
            if k and (target_l in str(k).lower() or str(k).lower() in target_l):
                return v
        return 0

    for reason, row in reason_rows_map.items():
        lq = _match_reason(reason, last_reason)
        tq = _match_reason(reason, this_reason)
        ws.cell(row=row, column=12, value=reason)
        ws.cell(row=row, column=13, value=lq)
        ws.cell(row=row, column=14, value=lq / last_reason_total).number_format = "0.00%"
        ws.cell(row=row, column=15, value=tq)
        ws.cell(row=row, column=16, value=tq / this_reason_total).number_format = "0.00%"
        top_style = ""
        for k, v in this_r["by_reason_top_style"].items():
            if reason.lower() in str(k).lower() or str(k).lower() in reason.lower():
                top_style = v
                break
        ws.cell(row=row, column=17, value=top_style)

    # ---------------- 销量 TOP 10 (B30:F39) ----------------
    for i, (_, rd) in enumerate(top10_sales.iterrows()):
        r = 30 + i
        ws.cell(row=r, column=2, value=rd["style"])
        ws.cell(row=r, column=3, value=int(rd["sales"]))
        ws.cell(row=r, column=4, value=int(rd["return_qty"]))
        ws.cell(row=r, column=5, value=rd["return_rate"]).number_format = "0.00%"
        ws.cell(row=r, column=6, value="")
    for i in range(len(top10_sales), 10):
        r = 30 + i
        for col in range(2, 7):
            ws.cell(row=r, column=col, value="")

    # ---------------- 退货率 TOP 10 (L30:P39) ----------------
    for i, (_, rd) in enumerate(top10_return.iterrows()):
        r = 30 + i
        ws.cell(row=r, column=12, value=rd["style"])
        ws.cell(row=r, column=13, value=int(rd["sales"]))
        ws.cell(row=r, column=14, value=int(rd["return_qty"]))
        ws.cell(row=r, column=15, value=rd["return_rate"]).number_format = "0.00%"
        ws.cell(row=r, column=16, value="")
    for i in range(len(top10_return), 10):
        r = 30 + i
        for col in range(12, 17):
            ws.cell(row=r, column=col, value="")

    # ---------------- 🌟 近30天新品 TOP 10 (B43:J52) ----------------
    # 列布局：B=款式 C=SKU F=销量 G=退货量 J=退货率
    for i in range(10):
        r = 43 + i
        if i < len(new_products_top10):
            rd = new_products_top10.iloc[i]
            ws.cell(row=r, column=2, value=rd["style"])
            ws.cell(row=r, column=3, value=rd["sku"])
            ws.cell(row=r, column=6, value=int(rd["sales"]))
            ws.cell(row=r, column=7, value=int(rd["return_qty"]))
            ws.cell(row=r, column=10, value=rd["return_rate"]).number_format = "0.00%"
        else:
            for col in [2, 3, 6, 7, 10]:
                ws.cell(row=r, column=col, value="")

    # ---------------- 🚨 问题款式 & 供应商 (L43:O46) ----------------
    # 列布局：L=供应商 M=问题款数 N=占比 O=代表款式
    for i in range(4):  # 模板里只有 4 行（43-46）
        r = 43 + i
        if i < len(supplier_problem):
            rd = supplier_problem.iloc[i]
            ws.cell(row=r, column=12, value=rd["supplier"])
            ws.cell(row=r, column=13, value=int(rd["problem_count"]))
            ws.cell(row=r, column=14, value=rd["ratio"]).number_format = "0.00%"
            ws.cell(row=r, column=15, value=rd["top_style"])
        else:
            for col in [12, 13, 14, 15]:
                ws.cell(row=r, column=col, value="")

    # ---------------- 输出 ----------------
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 9 部分：主入口 build_report                                   ║
# ╚════════════════════════════════════════════════════════════════════╝

def build_report(orders_file, returns_file, catalog_file,
                 template_bytes: bytes,
                 time_window_mode: str,
                 influencer_rule: str,
                 traffic_metrics: dict,
                 exclude_request_canceled: bool = True,
                 custom_last_label: str = None,
                 custom_this_label: str = None):
    """主入口：返回 (excel_bytes, summary_dict)"""

    influencer_col = INFLUENCER_RULE_MAP[influencer_rule]

    # 1. 读取
    orders_df  = _read_any(orders_file)
    returns_df = _read_any(returns_file)
    catalog_df = _read_any(catalog_file)

    if orders_df.empty:
        raise ValueError("订单数据为空")

    # 2. 切时间窗
    windows = split_time_windows(orders_df, returns_df, time_window_mode)
    last_label = custom_last_label or windows["last_label"]
    this_label = custom_this_label or windows["this_label"]

    # 3. 处理订单 / 退货
    last_o = process_orders(windows["last_orders"], influencer_col)
    this_o = process_orders(windows["this_orders"], influencer_col)
    last_r = process_returns(windows["last_returns"], exclude_request_canceled=exclude_request_canceled)
    this_r = process_returns(windows["this_returns"], exclude_request_canceled=exclude_request_canceled)

    # 4. 处理图册
    catalog = process_catalog(catalog_df)

    # 5. TOP 10 销量 / 退货率
    top10_sales = this_o["style_sales"].head(10).copy()
    top10_sales = top10_sales.merge(this_r["by_style"], on="style", how="left").fillna({"return_qty": 0})
    top10_sales["return_rate"] = top10_sales.apply(
        lambda r: r["return_qty"] / r["sales"] if r["sales"] > 0 else 0, axis=1
    )
    top10_sales = top10_sales[["style", "sales", "return_qty", "return_rate"]]

    style_full = this_o["style_sales"].merge(this_r["by_style"], on="style", how="left").fillna({"return_qty": 0})
    style_full = style_full[style_full["sales"] >= 10].copy()
    style_full["return_rate"] = style_full.apply(
        lambda r: r["return_qty"] / r["sales"] if r["sales"] > 0 else 0, axis=1
    )
    top10_return = style_full.sort_values("return_rate", ascending=False).head(10)[
        ["style", "sales", "return_qty", "return_rate"]
    ]

    # 6. 新品 + 供应商分析（基于本周）
    new_products = build_new_products_top10(this_o, this_r, catalog, windows["ref_date"])
    supplier_problem = build_supplier_problem(this_o, this_r, catalog)

    # 7. 写模板
    output_bytes = _write_template(
        template_bytes=template_bytes,
        last_label=last_label, this_label=this_label,
        last_o=last_o, this_o=this_o,
        last_r=last_r, this_r=this_r,
        top10_sales=top10_sales, top10_return=top10_return,
        new_products_top10=new_products,
        supplier_problem=supplier_problem,
        traffic_metrics=traffic_metrics,
    )

    # 8. 摘要
    return_rate_last = last_r["total_qty"] / last_o["sku_sold"] if last_o["sku_sold"] else 0
    return_rate_this = this_r["total_qty"] / this_o["sku_sold"] if this_o["sku_sold"] else 0
    base_keys = ["orders", "sku_sold", "gmv", "aov", "asp", "attach_rate", "influencer_orders"]

    summary = {
        "last": {**{k: last_o[k] for k in base_keys},
                 "return_rate": return_rate_last,
                 "total_qty": last_r["total_qty"],
                 "requested_qty": last_r["requested_qty"],
                 "request_canceled_qty": last_r["request_canceled_qty"],
                 "label": last_label},
        "this": {**{k: this_o[k] for k in base_keys},
                 "return_rate": return_rate_this,
                 "total_qty": this_r["total_qty"],
                 "requested_qty": this_r["requested_qty"],
                 "request_canceled_qty": this_r["request_canceled_qty"],
                 "label": this_label},
        "wow": {
            "orders":      _wow(this_o["orders"], last_o["orders"]),
            "gmv":         _wow(this_o["gmv"],    last_o["gmv"]),
            "aov":         _wow(this_o["aov"],    last_o["aov"]),
            "return_rate": return_rate_this - return_rate_last,
        },
        "top10_sales": top10_sales.assign(
            return_rate=top10_sales["return_rate"].apply(lambda x: f"{x:.1%}")
        ),
        "top10_return": top10_return.assign(
            return_rate=top10_return["return_rate"].apply(lambda x: f"{x:.1%}")
        ),
        "new_products": new_products.assign(
            return_rate=new_products["return_rate"].apply(lambda x: f"{x:.1%}")
        ) if not new_products.empty else new_products,
        "supplier_problem": supplier_problem.assign(
            ratio=supplier_problem["ratio"].apply(lambda x: f"{x:.1%}")
        ) if not supplier_problem.empty else supplier_problem,
        "windows": {
            "t_min": windows["t_min"],
            "t_max": windows["t_max"],
            "ref_date": windows["ref_date"],
        },
    }
    return output_bytes, summary


# ╔════════════════════════════════════════════════════════════════════╗
# ║   第 10 部分：Streamlit UI                                         ║
# ╚════════════════════════════════════════════════════════════════════╝

st.set_page_config(page_title="NailVesta 周报生成器", page_icon="💅", layout="wide")
st.title("💅 NailVesta 中台运营周报生成器 v2")
st.caption("上传 28 天数据 + 产品图册 → 自动切时间窗 + 识别新品/供应商 → 生成 WoW 周报")

# -------- 侧边栏 --------
with st.sidebar:
    st.header("⚙️ 参数")

    st.subheader("🕐 时间窗模式")
    time_window_mode_label = st.radio(
        "选择对比方式",
        ["28天对半切（前14 vs 后14）", "自然周对比（最近周 vs 上一周）"],
        index=0,
        help="28天对半切：数据完整、样本大；自然周对比：跟人类周报直觉吻合",
    )
    time_window_mode = "halve" if "对半切" in time_window_mode_label else "weekly"

    st.divider()
    st.subheader("周次标签覆盖（可选）")
    st.caption("不填则按时间窗自动生成")
    custom_last = st.text_input("上周标签", value="", help="如 4.6-4.19，留空则自动")
    custom_this = st.text_input("本周标签", value="", help="如 4.20-5.3，留空则自动")

    st.divider()
    st.subheader("达人单识别")
    influencer_rule = st.selectbox(
        "判定为达人单的条件",
        list(INFLUENCER_RULE_MAP.keys()),
        index=0,
    )

    st.divider()
    st.subheader("退货过滤")
    exclude_request_canceled = st.checkbox(
        "排除「Request Canceled」(买家主动撤销)",
        value=True,
        help="买家撤销不计入退货指标，但仍计入退货原因。客服拒绝（Refund rejected）保留。",
    )

    st.divider()
    st.subheader("可选 - TikTok 流量数据")
    st.caption("CTR / CVR / 加购率等，TikTok 后台导出后手填")
    traffic_metrics = {
        "ctr_last":       st.text_input("上周 CTR (%)"),
        "ctr_this":       st.text_input("本周 CTR (%)"),
        "cvr_last":       st.text_input("上周 CVR (%)"),
        "cvr_this":       st.text_input("本周 CVR (%)"),
        "cart_conv_last": st.text_input("上周 Cart Conv (%)"),
        "cart_conv_this": st.text_input("本周 Cart Conv (%)"),
        "atc_last":       st.text_input("上周 Add-To-Cart (%)"),
        "atc_this":       st.text_input("本周 Add-To-Cart (%)"),
    }

# -------- 主区域：上传 --------
st.subheader("📂 上传 3 个文件")
st.caption("订单和退货是 28 天数据；产品图册用于识别供应商和新品（≤30 天）")

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("**1. 订单数据（28天）**")
    orders_file = st.file_uploader("All_order csv/xlsx", type=["csv", "xlsx"], key="orders")
with col2:
    st.markdown("**2. 退货数据（28天）**")
    returns_file = st.file_uploader("Return_Refund_Orders xlsx", type=["csv", "xlsx"], key="returns")
with col3:
    st.markdown("**3. 产品图册**")
    DEFAULT_CATALOG = Path(__file__).parent / "NailVesta_产品图册.csv"
    catalog_file = st.file_uploader(
        "NailVesta_产品图册 csv（可选）",
        type=["csv", "xlsx"],
        key="catalog",
        help=f"留空则用程序自带的 {DEFAULT_CATALOG.name}（如果有）",
    )

DEFAULT_TEMPLATE = Path(__file__).parent / "NailVesta_中台运营周报_模板.xlsx"
with st.expander("（可选）自定义周报模板"):
    template_file = st.file_uploader(
        "上传自定义模板（不传就用程序自带的）",
        type=["xlsx"], key="template",
    )

st.divider()

# -------- 生成按钮 --------
if st.button("🚀 生成周报", type="primary", use_container_width=True):
    if not orders_file:
        st.error("⚠️ 请上传订单文件")
        st.stop()

    # 模板
    if 'template_file' in dir() and template_file is not None:
        template_bytes = template_file.read()
    else:
        if not DEFAULT_TEMPLATE.exists():
            st.error(f"⚠️ 默认模板不存在：{DEFAULT_TEMPLATE}")
            st.stop()
        template_bytes = DEFAULT_TEMPLATE.read_bytes()

    # 图册：如果没上传就用默认
    if catalog_file is None:
        if DEFAULT_CATALOG.exists():
            catalog_file_obj = open(DEFAULT_CATALOG, "rb")
            class _F:
                def __init__(self, raw, name):
                    self._raw = raw
                    self.name = name
                    self._bytes = raw.read()
                def seek(self, p): pass
                def read(self): return self._bytes
            with open(DEFAULT_CATALOG, "rb") as fh:
                catalog_bytes = fh.read()
            catalog_file = io.BytesIO(catalog_bytes)
            catalog_file.name = DEFAULT_CATALOG.name
        else:
            st.warning("⚠️ 未上传图册，且程序目录无默认图册，将无法识别供应商和新品")

    with st.spinner("正在跑数据... ☕"):
        try:
            output_bytes, summary = build_report(
                orders_file=orders_file,
                returns_file=returns_file,
                catalog_file=catalog_file,
                template_bytes=template_bytes,
                time_window_mode=time_window_mode,
                influencer_rule=influencer_rule,
                exclude_request_canceled=exclude_request_canceled,
                traffic_metrics=traffic_metrics,
                custom_last_label=custom_last or None,
                custom_this_label=custom_this or None,
            )
        except Exception as e:
            import traceback
            st.error(f"❌ 生成失败：{e}")
            st.code(traceback.format_exc())
            st.stop()

    st.success("✅ 周报已生成！")

    # 时间窗信息
    w = summary["windows"]
    st.info(
        f"📅 数据时间范围：{w['t_min'].strftime('%Y-%m-%d')} ~ {w['t_max'].strftime('%Y-%m-%d')}  "
        f"|  上周：**{summary['last']['label']}**  vs  本周：**{summary['this']['label']}**  "
        f"|  新品基准日：{w['ref_date'].strftime('%Y-%m-%d')}（前 30 天上架的算新品）"
    )

    # 看板：核心指标
    st.subheader("📊 本周快速看板")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("订单数（剔除达人）", f"{summary['this']['orders']:,}",
              f"{summary['wow']['orders']:+.1%}")
    c2.metric("总GMV ($)", f"{summary['this']['gmv']:,.0f}",
              f"{summary['wow']['gmv']:+.1%}")
    c3.metric("AOV ($)", f"{summary['this']['aov']:.2f}",
              f"{summary['wow']['aov']:+.1%}")
    c4.metric("退货率", f"{summary['this']['return_rate']:.2%}",
              f"{summary['wow']['return_rate']:+.1%}", delta_color="inverse")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("达人单（上周）", f"{summary['last']['influencer_orders']:,}")
    c6.metric("达人单（本周）", f"{summary['this']['influencer_orders']:,}")
    c7.metric("SKU 销量", f"{summary['this']['sku_sold']:,}")
    c8.metric("连带率", f"{summary['this']['attach_rate']:.2f}")

    # 退货拆解
    st.markdown("**📦 退货拆解** —— 实际退货量 = 申请总量 − 买家撤销")
    c9, c10, c11, c12 = st.columns(4)
    c9.metric("申请退货总量", f"{summary['this']['requested_qty']:,}")
    c10.metric("买家撤销", f"{summary['this']['request_canceled_qty']:,}")
    c11.metric("实际退货量", f"{summary['this']['total_qty']:,}",
               help="含完成退款 + 客服拒绝（后者反映客服质量）")
    cancel_rate = (summary['this']['request_canceled_qty'] / summary['this']['requested_qty']
                    if summary['this']['requested_qty'] else 0)
    c12.metric("买家撤销率", f"{cancel_rate:.1%}",
               help="数字越高 = 客服安抚越成功 ✨")

    # TOP 10 表
    st.subheader("🏆 销量 TOP 10 款式（本周）")
    st.dataframe(summary["top10_sales"], use_container_width=True, hide_index=True)

    st.subheader("⚠️ 退货率 TOP 10 款式（本周，销量≥10）")
    st.dataframe(summary["top10_return"], use_container_width=True, hide_index=True)

    # 新品 TOP 10
    st.subheader("🌟 近30天新品销量 TOP 10")
    if not summary["new_products"].empty:
        st.dataframe(summary["new_products"], use_container_width=True, hide_index=True)
    else:
        st.caption("（暂无新品数据，确认产品图册已上传且包含上架时间字段）")

    # 供应商问题
    st.subheader("🚨 问题款式 & 供应商（退货率≥10% 且销量≥10）")
    if not summary["supplier_problem"].empty:
        st.dataframe(summary["supplier_problem"], use_container_width=True, hide_index=True)
    else:
        st.caption("（本周没有问题款式，或图册未提供供应商字段）")

    # 下载
    fname = f"NailVesta_周报_{summary['this']['label'].replace('/', '-')}.xlsx"
    st.download_button(
        "📥 下载完整周报 Excel",
        data=output_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.info("👆 上传订单 + 退货 + 产品图册后点「生成周报」就行啦~")

    with st.expander("📖 使用说明"):
        st.markdown(
            """
            **每周操作流程**

            1. TikTok Shop 后台导出**最近 28 天**的订单 + 退货数据。
            2. 上传订单、退货、产品图册三个文件（图册不传则用程序自带的）。
            3. 在左侧选「时间窗模式」：28天对半切 / 自然周对比。
            4. 点击「生成周报」即可下载。

            **时间窗对比**

            | 模式 | 上周 | 本周 | 优势 |
            |---|---|---|---|
            | 28天对半切（默认） | 前 14 天 | 后 14 天 | 数据完整，样本大 |
            | 自然周对比 | 上一周（一-日） | 最近完整周 | 跟人类周报直觉一致 |

            **达人单剔除**：默认按「SKU Subtotal After Discount = 0」。

            **退货数据双层处理**

            | 指标 | 含 Refund rejected | 含 Request Canceled |
            |---|---|---|
            | 退货量 / 退货率 / 金额 / TOP10 | ✅ 含 | ❌ 不含 |
            | 退货原因统计 | ✅ 含 | ✅ 含 |

            **新品识别**：图册中「上架时间」≤ 30 天的款 = 新品

            **供应商问题**：本周退货率 ≥ 10% 且销量 ≥ 10 的款 = 问题款，按图册「厂家」字段聚合
            """
        )
