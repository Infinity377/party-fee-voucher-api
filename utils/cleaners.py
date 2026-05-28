from __future__ import annotations

from pathlib import Path
from typing import Any
import re
import pandas as pd


def normalize_cell(value: Any) -> str:
    """
    将 Excel 单元格值统一转为去空格字符串。
    空值返回空字符串。
    """
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_amount(value: Any) -> float | None:
    """
    将金额字段转为 float。
    支持逗号、空格、字符串金额。
    空值返回 None。
    """
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text == "":
        return None

    text = text.replace(",", "").replace("，", "")
    text = text.replace("￥", "").replace("¥", "")

    try:
        return float(text)
    except ValueError:
        return None


def normalize_date(value: Any) -> str:
    """
    将日期或日期时间标准化为 YYYY-MM-DD。
    无法识别时返回原文本。
    """
    if pd.isna(value):
        return ""

    try:
        dt = pd.to_datetime(value)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return str(value).strip()


def month_end_date(value: Any) -> str:
    """
    根据交易日期返回所在月份月末日期，格式 YYYY-MM-DD。
    """
    if pd.isna(value):
        return ""

    try:
        dt = pd.to_datetime(value)
        return (dt + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")
    except Exception:
        return ""


def extract_flow_no(*texts: Any) -> str:
    """
    从多个文本字段中提取 OA 流程编号。
    支持“报销款 DW0226030001”和“报销款DW0226030001”两种写法。
    """
    combined = " ".join(normalize_cell(t) for t in texts if normalize_cell(t))
    match = re.search(r"DW\d{8,}", combined, flags=re.IGNORECASE)
    if match:
        return match.group(0).upper()
    return ""


def split_keywords(value: Any) -> list[str]:
    """
    将规则表中的关键词字段拆成列表。
    支持 /、中文顿号、逗号、分号、换行。
    """
    text = normalize_cell(value)
    if text == "" or text == "无":
        return []

    parts = re.split(r"[\/、,，;；\n\r]+", text)
    return [p.strip() for p in parts if p.strip()]


def contains_any(text: str, keywords: list[str]) -> bool:
    """
    判断 text 中是否包含任意关键词。
    """
    if not keywords:
        return False
    return any(keyword in text for keyword in keywords)


def clean_bank_flow(file_path: Path) -> pd.DataFrame:
    """
    清洗银行流水表，输出标准字段。
    """
    raw = pd.read_excel(file_path, sheet_name=0, header=1)
    raw = raw.dropna(how="all").reset_index(drop=True)

    rows: list[dict[str, Any]] = []

    for row_no, (_, row) in enumerate(raw.iterrows(), start=1):
        out_amount = normalize_amount(row.get("转出金额"))
        in_amount = normalize_amount(row.get("转入金额"))
        debit_credit = normalize_cell(row.get("借贷标志"))

        if debit_credit == "贷" or (in_amount is not None and in_amount > 0):
            direction = "收入"
            amount = in_amount
        elif debit_credit == "借" or (out_amount is not None and out_amount > 0):
            direction = "支出"
            amount = out_amount
        else:
            direction = "未知"
            amount = None

        transaction_date = normalize_date(row.get("交易时间") or row.get("入账日期"))
        counterparty = normalize_cell(row.get("对方单位"))
        purpose = normalize_cell(row.get("用途"))
        bank_summary = normalize_cell(row.get("摘要"))
        postscript = normalize_cell(row.get("附言"))
        receipt_info = normalize_cell(row.get("回单个性化信息"))

        combined_text = " ".join(
            x for x in [counterparty, purpose, bank_summary, postscript, receipt_info] if x
        )

        oa_flow_no = extract_flow_no(purpose, bank_summary, postscript, receipt_info)

        rows.append(
            {
                "flow_index": row_no,
                "transaction_datetime": normalize_cell(row.get("交易时间")),
                "transaction_date": transaction_date,
                "voucher_no_raw": normalize_cell(row.get("凭证号")),
                "debit_credit": debit_credit,
                "out_amount": out_amount,
                "in_amount": in_amount,
                "amount": amount,
                "direction": direction,
                "counterparty": counterparty,
                "purpose": purpose,
                "bank_summary": bank_summary,
                "postscript": postscript,
                "receipt_info": receipt_info,
                "combined_text": combined_text,
                "oa_flow_no": oa_flow_no,
                "voucher_date": month_end_date(row.get("交易时间") or row.get("入账日期")),
                "raw_row_no": row_no,
            }
        )

    return pd.DataFrame(rows)


def clean_oa_flow(file_path: Path) -> pd.DataFrame:
    """
    清洗 OA 划款流程表，输出标准字段。
    """
    raw = pd.read_excel(file_path, sheet_name=0, header=0)
    raw = raw.dropna(how="all").reset_index(drop=True)

    rows: list[dict[str, Any]] = []

    for row_no, (_, row) in enumerate(raw.iterrows(), start=1):
        title = normalize_cell(row.get("申请标题"))

        rows.append(
            {
                "oa_flow_no": normalize_cell(row.get("流程编号")).upper(),
                "submit_date": normalize_date(row.get("提交日期")),
                "application_title": title,
                "oa_amount": normalize_amount(row.get("金额小写")),
                "title_clean_base": title,
                "raw_row_no": row_no,
            }
        )

    return pd.DataFrame(rows)


def clean_subject_table(file_path: Path) -> pd.DataFrame:
    """
    清洗会计科目表。
    当前要求该文件已经删除前 6 行，第一行为字段名。
    """
    raw = pd.read_excel(file_path, sheet_name=0, header=0)
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    rows: list[dict[str, Any]] = []

    for _, row in raw.iterrows():
        subject_code = normalize_cell(row.get("科目编码"))

        if subject_code == "":
            continue

        rows.append(
            {
                "subject_code": subject_code,
                "subject_name": normalize_cell(row.get("科目名称")),
                "subject_type": normalize_cell(row.get("科目类型")),
                "cash_category": normalize_cell(row.get("现金分类")),
                "direction": normalize_cell(row.get("方向")),
                "auxiliary_accounting": normalize_cell(row.get("辅助核算")),
                "is_leaf": normalize_cell(row.get("末级")),
                "unit": normalize_cell(row.get("计量单位")),
            }
        )

    return pd.DataFrame(rows)


def clean_member_status_table(file_path: Path) -> pd.DataFrame:
    """
    清洗党员离退休情况表。

    预期字段：
    姓名、状态、对应收入科目编码、对应收入科目名称
    """
    raw = pd.read_excel(file_path, sheet_name=0, header=0)
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    rows: list[dict[str, Any]] = []

    for _, row in raw.iterrows():
        name = normalize_cell(row.get("姓名"))
        if name == "":
            continue

        rows.append(
            {
                "member_name": name,
                "member_status": normalize_cell(row.get("状态")),
                "income_subject_code": normalize_cell(row.get("对应收入科目编码")),
                "income_subject_name": normalize_cell(row.get("对应收入科目名称")),
            }
        )

    return pd.DataFrame(rows)


def clean_rule_mapping_table(file_path: Path) -> pd.DataFrame:
    """
    清洗党费业务映射规则表。
    """
    raw = pd.read_excel(file_path, sheet_name=0, header=0)
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    rows: list[dict[str, Any]] = []

    for _, row in raw.iterrows():
        rule_id = normalize_cell(row.get("规则编号"))
        if rule_id == "":
            continue

        priority = normalize_amount(row.get("优先级"))
        if priority is None:
            priority = 0

        rows.append(
            {
                "rule_id": rule_id,
                "rule_name": normalize_cell(row.get("规则名称")),
                "enabled": normalize_cell(row.get("启用状态")),
                "priority": int(priority),
                "direction": normalize_cell(row.get("收支方向")),
                "target_object": normalize_cell(row.get("适用对象")),
                "need_oa": normalize_cell(row.get("是否需要OA匹配")),
                "need_member_status": normalize_cell(row.get("是否需要党员状态表")),
                "match_fields": normalize_cell(row.get("匹配字段")),
                "match_method": normalize_cell(row.get("匹配方式")),
                "keywords": normalize_cell(row.get("关键词")),
                "exclude_keywords": normalize_cell(row.get("排除关键词")),
                "keyword_list": split_keywords(row.get("关键词")),
                "exclude_keyword_list": split_keywords(row.get("排除关键词")),
                "subject_code": normalize_cell(row.get("推荐科目编码")),
                "subject_name": normalize_cell(row.get("推荐科目名称")),
                "voucher_summary_template": normalize_cell(row.get("凭证摘要模板")),
                "period_extract_rule": normalize_cell(row.get("期间提取规则")),
                "summary_clean_rule": normalize_cell(row.get("摘要清洗规则")),
                "ledger_tag": normalize_cell(row.get("台账标签")),
                "confidence": normalize_cell(row.get("置信度")),
                "exception_handling": normalize_cell(row.get("异常处理")),
            }
        )

    return pd.DataFrame(rows)


def dataframe_to_records(df: pd.DataFrame, max_rows: int = 20) -> list[dict[str, Any]]:
    """
    将 DataFrame 转成适合 JSON 返回的 records。
    默认最多返回 20 行，避免响应过长。
    """
    records: list[dict[str, Any]] = []
    for _, row in df.head(max_rows).iterrows():
        records.append({str(col): normalize_json_value(row[col]) for col in df.columns})
    return records


def normalize_json_value(value: Any) -> Any:
    """
    转换 JSON 友好的值。
    需要优先处理 list / dict，否则 pd.isna(list) 会触发
    The truth value of an empty array is ambiguous.
    """
    if isinstance(value, list):
        return [normalize_json_value(item) for item in value]

    if isinstance(value, dict):
        return {
            str(key): normalize_json_value(val)
            for key, val in value.items()
        }

    if isinstance(value, tuple):
        return [normalize_json_value(item) for item in value]

    if pd.isna(value):
        return ""

    if hasattr(value, "isoformat"):
        return value.isoformat()

    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass

    return value