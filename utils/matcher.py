from __future__ import annotations

from typing import Any
import re
import pandas as pd

from utils.cleaners import contains_any, normalize_cell


def series_to_str_dict(row: pd.Series) -> dict[str, Any]:
    """
    将 pandas Series 转成 dict[str, Any]，避免 Pylance 将 to_dict() 识别为 dict[Hashable, Any]。
    """
    return {str(key): value for key, value in row.to_dict().items()}


def is_company_or_org(name: str) -> bool:
    """
    粗略判断对方单位是否为公司/机构。
    Demo 阶段使用关键词判断。
    """
    org_keywords = [
        "公司", "基金", "管理", "投资", "发展", "有限", "集团",
        "银行", "党支部", "党总支", "党委", "中心", "协会"
    ]
    return any(keyword in name for keyword in org_keywords)


def extract_interest_period_text(text: str, fallback_date: str = "") -> str:
    """
    专门提取银行利息期间。

    优先级：
    1. 起息日期 + 止息日期，例如：
       起息日期:2025-12-21 止息日期:2026-03-20
       -> 2025-12-21到2026-03-20
    2. 根据止息日期判断季度，例如 2026-03-20 -> 2026年一季度
    3. 根据交易日期判断季度
    4. 返回空字符串
    """
    text = normalize_cell(text)

    start_match = re.search(r"起息日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})", text)
    end_match = re.search(r"止息日期[:：]\s*(\d{4}-\d{1,2}-\d{1,2})", text)

    if start_match and end_match:
        start_date = start_match.group(1)
        end_date = end_match.group(1)
        return f"{start_date}到{end_date}"

    # 如果只有止息日期，则按止息日期判断季度
    if end_match:
        end_date = end_match.group(1)
        try:
            dt = pd.to_datetime(end_date)
            quarter_map = {
                1: "一季度",
                2: "二季度",
                3: "三季度",
                4: "四季度",
            }
            return f"{dt.year}年{quarter_map[dt.quarter]}"
        except Exception:
            pass

    # 如果没有起止息日期，则尝试按交易日期判断季度
    if fallback_date:
        try:
            dt = pd.to_datetime(fallback_date)
            quarter_map = {
                1: "一季度",
                2: "二季度",
                3: "三季度",
                4: "四季度",
            }
            return f"{dt.year}年{quarter_map[dt.quarter]}"
        except Exception:
            pass

    return ""


def extract_period_text(text: str, fallback_date: str = "") -> str:
    """
    从文本中提取普通业务期间。
    Demo 版支持：
    - 2026年3月
    - 3月
    - 2504-2604
    若无法提取，使用交易日期所在月份。
    """
    text = normalize_cell(text)

    match_full = re.search(r"(20\d{2})年\s*(\d{1,2})月", text)
    if match_full:
        return f"{match_full.group(1)}年{int(match_full.group(2))}月"

    match_month = re.search(r"(?<!\d)(\d{1,2})月", text)
    if match_month:
        return f"{int(match_month.group(1))}月"

    match_range = re.search(r"(\d{2})(\d{2})[-至到](\d{2})(\d{2})", text)
    if match_range:
        start_y = "20" + match_range.group(1)
        start_m = int(match_range.group(2))
        end_y = "20" + match_range.group(3)
        end_m = int(match_range.group(4))
        return f"{start_y}年{start_m}月至{end_y}年{end_m}月"

    if fallback_date:
        parts = fallback_date.split("-")
        if len(parts) >= 2:
            return f"{int(parts[1])}月"

    return ""


def should_use_backpay_word(period_text: str, transaction_date: str) -> bool:
    """
    判断摘要里用“缴纳”还是“补缴”。
    Demo 规则：
    - 期间中出现“至”或跨度，使用补缴；
    - 否则默认缴纳。
    """
    if "至" in period_text or "到" in period_text or "-" in period_text:
        return True
    return False


def balance_parentheses(text: str) -> str:
    """
    修复中文/英文括号不成对问题。
    """
    left_cn = text.count("（")
    right_cn = text.count("）")
    if left_cn > right_cn:
        text += "）" * (left_cn - right_cn)

    left_en = text.count("(")
    right_en = text.count(")")
    if left_en > right_en:
        text += ")" * (left_en - right_en)

    return text


def clean_title_for_summary(title: str) -> str:
    """
    清洗 OA 标题，生成凭证摘要使用。

    注意：
    1. “请示”只用于判断是否加“预付款”，不直接保留在摘要中。
    2. 不直接删除“费用”，避免破坏原始业务含义。
    3. 保留括号内容，如“（第五期）”。
    """
    text = normalize_cell(title)

    # 先处理常见首部套话
    prefix_patterns = [
        r"^关于",
        r"^报销",
        r"^申请",
        r"^请示",
    ]
    for pattern in prefix_patterns:
        text = re.sub(pattern, "", text)

    # 处理常见尾部套话
    suffix_patterns = [
        r"的请示$",
        r"请示$",
        r"的申请$",
        r"申请$",
        r"费用报销$",
        r"报销$",
    ]
    for pattern in suffix_patterns:
        text = re.sub(pattern, "", text)

    # 删除中间无意义词，但不要删除“费用”
    remove_words = [
        "报销款",
        "费用报销",
    ]
    for word in remove_words:
        text = text.replace(word, "")

    text = re.sub(r"\s+", "", text)
    text = text.strip("：:，,。；; ")

    text = balance_parentheses(text)

    return text


def build_expense_summary(oa_flow_no: str, application_title: str, voucher_summary_template: str = "") -> str:
    """
    根据 OA 标题和摘要模板生成支出凭证摘要。

    规则：
    1. 基本格式：支付{流程编号}{申请标题清洗后}
    2. 如果模板以“费用”结尾，或者标题中有“活动/培训/教育/会议”等费用性质词，则补“费用”
    3. 如果原始 OA 标题包含“请示”，则追加“预付款”
    4. 避免“费用费用”“预付款预付款”
    """
    raw_title = normalize_cell(application_title)
    template = normalize_cell(voucher_summary_template)

    title_clean = clean_title_for_summary(raw_title)

    summary = f"支付{oa_flow_no}{title_clean}"

    need_fee_word = False

    if template.endswith("费用"):
        need_fee_word = True

    fee_keywords = [
        "活动", "培训", "教育", "会议", "讲座", "党建共建",
        "主题党日", "革命传统教育", "学习培训"
    ]
    if any(keyword in raw_title for keyword in fee_keywords):
        need_fee_word = True

    if need_fee_word and not summary.endswith("费") and not summary.endswith("费用"):
        summary += "费用"

    if "请示" in raw_title and not summary.endswith("预付款"):
        summary += "预付款"

    summary = summary.replace("费用费用", "费用")
    summary = summary.replace("预付款预付款", "预付款")
    summary = balance_parentheses(summary)

    return summary


def build_oa_lookup(oa_df: pd.DataFrame) -> dict[str, dict[str, Any]]:
    """
    将 OA 表转换为按流程编号索引的字典。
    """
    lookup: dict[str, dict[str, Any]] = {}

    for _, row in oa_df.iterrows():
        row_dict = series_to_str_dict(row)
        flow_no = normalize_cell(row_dict.get("oa_flow_no")).upper()
        if flow_no:
            lookup[flow_no] = row_dict

    return lookup


def build_member_lookup(member_df: pd.DataFrame) -> dict[str, dict[str, Any]]:
    """
    将党员状态表转换为按姓名索引的字典。
    """
    lookup: dict[str, dict[str, Any]] = {}

    for _, row in member_df.iterrows():
        row_dict = series_to_str_dict(row)
        name = normalize_cell(row_dict.get("member_name"))
        if name:
            lookup[name] = row_dict

    return lookup


def validate_subject(subject_df: pd.DataFrame, subject_code: str) -> dict[str, Any]:
    """
    校验科目是否存在、是否末级。
    """
    subject_code = normalize_cell(subject_code)
    matched = subject_df[subject_df["subject_code"].astype(str) == subject_code]

    if matched.empty:
        return {
            "valid": False,
            "subject_code": subject_code,
            "subject_name": "",
            "subject_type": "",
            "direction": "",
            "is_leaf": "",
            "error": "科目不存在于会计科目表"
        }

    row = series_to_str_dict(matched.iloc[0])
    is_leaf = normalize_cell(row.get("is_leaf"))

    if is_leaf != "是":
        return {
            "valid": False,
            "subject_code": subject_code,
            "subject_name": normalize_cell(row.get("subject_name")),
            "subject_type": normalize_cell(row.get("subject_type")),
            "direction": normalize_cell(row.get("direction")),
            "is_leaf": is_leaf,
            "error": "科目不是末级科目"
        }

    return {
        "valid": True,
        "subject_code": subject_code,
        "subject_name": normalize_cell(row.get("subject_name")),
        "subject_type": normalize_cell(row.get("subject_type")),
        "direction": normalize_cell(row.get("direction")),
        "is_leaf": is_leaf,
        "error": ""
    }


def match_income_rule(
    flow_row: pd.Series,
    member_lookup: dict[str, dict[str, Any]],
    subject_df: pd.DataFrame
) -> dict[str, Any]:
    """
    匹配收入类流水。
    """
    counterparty = normalize_cell(flow_row.get("counterparty"))
    combined_text = normalize_cell(flow_row.get("combined_text"))
    transaction_date = normalize_cell(flow_row.get("transaction_date"))

    # 1. 银行利息：优先使用起息日期 / 止息日期
    if "利息" in combined_text:
        subject_code = "400103"
        subject_check = validate_subject(subject_df, subject_code)
        interest_period = extract_interest_period_text(combined_text, transaction_date)

        if interest_period:
            summary = f"收到{interest_period}银行利息"
        else:
            summary = "收到银行利息"

        return {
            "match_status": "matched",
            "business_type": "银行利息收入",
            "rule_id": "R001",
            "subject_code": subject_code,
            "subject_name": subject_check["subject_name"],
            "summary": summary,
            "ledger_tag": "党费账户利息收入",
            "confidence": "高",
            "exception": "" if subject_check["valid"] else subject_check["error"]
        }

    # 2. 个人党员党费：优先看党员状态表
    if counterparty in member_lookup and "党费" in combined_text:
        member = member_lookup[counterparty]
        subject_code = normalize_cell(member.get("income_subject_code"))
        subject_name = normalize_cell(member.get("income_subject_name"))
        subject_check = validate_subject(subject_df, subject_code)

        period = extract_period_text(combined_text, transaction_date)
        verb = "补缴" if should_use_backpay_word(period, transaction_date) else "缴纳"

        return {
            "match_status": "matched",
            "business_type": f"个人{normalize_cell(member.get('member_status'))}党员党费",
            "rule_id": "R003/R004/R005",
            "subject_code": subject_code,
            "subject_name": subject_name or subject_check["subject_name"],
            "summary": f"收到{counterparty}{verb}{period}党费" if period else f"收到{counterparty}{verb}党费",
            "ledger_tag": "其他党员自行上缴",
            "confidence": "高",
            "exception": "" if subject_check["valid"] else subject_check["error"]
        }

    # 3. 公司/机构统一上缴党费
    if is_company_or_org(counterparty) and "党费" in combined_text:
        subject_code = "400101"
        subject_check = validate_subject(subject_df, subject_code)
        period = extract_period_text(combined_text, transaction_date)

        return {
            "match_status": "matched",
            "business_type": "公司统一上缴党费",
            "rule_id": "R002",
            "subject_code": subject_code,
            "subject_name": subject_check["subject_name"],
            "summary": f"收到{counterparty}划来{period}代收党费" if period else f"收到{counterparty}划来代收党费",
            "ledger_tag": "公司党员统一上缴",
            "confidence": "高",
            "exception": "" if subject_check["valid"] else subject_check["error"]
        }

    # 4. 看起来是个人党费，但党员状态表缺失
    if "党费" in combined_text and counterparty:
        return {
            "match_status": "exception",
            "business_type": "个人党费收入",
            "rule_id": "",
            "subject_code": "",
            "subject_name": "",
            "summary": "",
            "ledger_tag": "",
            "confidence": "低",
            "exception": f"收入流水疑似个人党费，但未在党员离退休情况表中找到：{counterparty}"
        }

    return {
        "match_status": "exception",
        "business_type": "收入类别无法判断",
        "rule_id": "",
        "subject_code": "",
        "subject_name": "",
        "summary": "",
        "ledger_tag": "",
        "confidence": "低",
        "exception": "收入流水无法匹配收入类规则"
    }


def match_expense_rule(
    flow_row: pd.Series,
    oa_lookup: dict[str, dict[str, Any]],
    rule_df: pd.DataFrame,
    subject_df: pd.DataFrame
) -> dict[str, Any]:
    """
    匹配支出类流水：先匹配 OA，再匹配业务规则表。
    """
    oa_flow_no = normalize_cell(flow_row.get("oa_flow_no")).upper()
    amount = flow_row.get("amount")

    if not oa_flow_no:
        return {
            "match_status": "exception",
            "business_type": "支出OA缺失",
            "oa_flow_no": "",
            "oa_match_status": "未匹配",
            "rule_id": "",
            "subject_code": "",
            "subject_name": "",
            "summary": "",
            "ledger_tag": "",
            "confidence": "低",
            "exception": "支出流水未提取到OA流程编号"
        }

    if oa_flow_no not in oa_lookup:
        return {
            "match_status": "exception",
            "business_type": "支出OA不存在",
            "oa_flow_no": oa_flow_no,
            "oa_match_status": "未匹配",
            "rule_id": "",
            "subject_code": "",
            "subject_name": "",
            "summary": "",
            "ledger_tag": "",
            "confidence": "低",
            "exception": f"支出流水中的流程编号 {oa_flow_no} 未在OA划款流程表中找到"
        }

    oa_row = oa_lookup[oa_flow_no]
    oa_amount = oa_row.get("oa_amount")

    if amount is None or oa_amount is None or abs(float(amount) - float(oa_amount)) > 0.005:
        return {
            "match_status": "exception",
            "business_type": "支出金额不一致",
            "oa_flow_no": oa_flow_no,
            "oa_match_status": "已匹配但金额不一致",
            "rule_id": "",
            "subject_code": "",
            "subject_name": "",
            "summary": "",
            "ledger_tag": "",
            "confidence": "低",
            "exception": f"流水金额 {amount} 与 OA 金额 {oa_amount} 不一致"
        }

    application_title = normalize_cell(oa_row.get("application_title"))
    combined_text = " ".join([
        application_title,
        normalize_cell(flow_row.get("combined_text"))
    ])

    expense_rules = rule_df[
        (rule_df["enabled"] == "启用") &
        (rule_df["direction"] == "支出")
    ].copy()

    matched_rules: list[dict[str, Any]] = []

    for _, rule in expense_rules.iterrows():
        rule_dict = series_to_str_dict(rule)
        keywords = rule_dict.get("keyword_list", [])
        exclude_keywords = rule_dict.get("exclude_keyword_list", [])

        if not isinstance(keywords, list):
            keywords = []
        if not isinstance(exclude_keywords, list):
            exclude_keywords = []

        if contains_any(combined_text, exclude_keywords):
            continue

        if contains_any(combined_text, keywords):
            matched_rules.append(rule_dict)

    if not matched_rules:
        return {
            "match_status": "exception",
            "business_type": "支出规则缺失",
            "oa_flow_no": oa_flow_no,
            "oa_match_status": "已匹配",
            "rule_id": "",
            "subject_code": "",
            "subject_name": "",
            "summary": "",
            "ledger_tag": "",
            "confidence": "低",
            "exception": "支出业务未匹配到业务映射规则表"
        }

    matched_rules = sorted(matched_rules, key=lambda x: int(x.get("priority", 0)), reverse=True)
    selected_rule = matched_rules[0]

    subject_code = normalize_cell(selected_rule.get("subject_code"))
    subject_check = validate_subject(subject_df, subject_code)

    summary = build_expense_summary(
        oa_flow_no=oa_flow_no,
        application_title=application_title,
        voucher_summary_template=normalize_cell(selected_rule.get("voucher_summary_template"))
    )

    return {
        "match_status": "matched" if subject_check["valid"] else "exception",
        "business_type": normalize_cell(selected_rule.get("rule_name")),
        "oa_flow_no": oa_flow_no,
        "oa_match_status": "已匹配",
        "rule_id": normalize_cell(selected_rule.get("rule_id")),
        "subject_code": subject_code,
        "subject_name": normalize_cell(selected_rule.get("subject_name")) or subject_check["subject_name"],
        "summary": summary,
        "ledger_tag": normalize_cell(selected_rule.get("ledger_tag")),
        "confidence": normalize_cell(selected_rule.get("confidence")),
        "exception": "" if subject_check["valid"] else subject_check["error"]
    }


def match_all_business(
    bank_df: pd.DataFrame,
    oa_df: pd.DataFrame,
    subject_df: pd.DataFrame,
    member_df: pd.DataFrame,
    rule_df: pd.DataFrame
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    """
    对全部流水进行业务匹配。
    返回：
    1. matched_df：带匹配结果的业务明细
    2. exceptions：异常清单
    """
    oa_lookup = build_oa_lookup(oa_df)
    member_lookup = build_member_lookup(member_df)

    matched_rows: list[dict[str, Any]] = []
    exceptions: list[dict[str, Any]] = []

    for _, flow_row in bank_df.iterrows():
        direction = normalize_cell(flow_row.get("direction"))

        if direction == "收入":
            match_result = match_income_rule(flow_row, member_lookup, subject_df)
        elif direction == "支出":
            match_result = match_expense_rule(flow_row, oa_lookup, rule_df, subject_df)
        else:
            match_result = {
                "match_status": "exception",
                "business_type": "流水方向未知",
                "rule_id": "",
                "subject_code": "",
                "subject_name": "",
                "summary": "",
                "ledger_tag": "",
                "confidence": "低",
                "exception": "无法根据借贷标志或金额判断收入/支出方向"
            }

        row_dict = series_to_str_dict(flow_row)
        row_dict.update(match_result)
        matched_rows.append(row_dict)

        if normalize_cell(match_result.get("match_status")) != "matched" or normalize_cell(match_result.get("exception")):
            exceptions.append({
                "flow_index": flow_row.get("flow_index"),
                "transaction_date": flow_row.get("transaction_date"),
                "direction": flow_row.get("direction"),
                "amount": flow_row.get("amount"),
                "counterparty": flow_row.get("counterparty"),
                "oa_flow_no": match_result.get("oa_flow_no", flow_row.get("oa_flow_no", "")),
                "exception": match_result.get("exception"),
                "suggestion": "请人工确认该笔业务，并补充党员状态表或业务映射规则表。"
            })

    matched_df = pd.DataFrame(matched_rows)
    return matched_df, exceptions