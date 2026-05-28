from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from utils.cleaners import normalize_cell
from utils.matcher import match_all_business
from utils.voucher_generator import (
    find_header_row as find_voucher_header_row,
    get_template_column_map as get_voucher_column_map,
    normalize_amount_for_excel,
)
from utils.ledger_generator import (
    find_ledger_header_row,
    get_ledger_column_map,
)


BANK_SUBJECT_CODE = "1002"
AI_REVIEW_MARKERS = ["AI建议补录", "AI补录", "待复核"]


# =========================
# 基础规范化函数
# =========================

def normalize_amount(value: Any) -> float:
    """
    统一金额格式，便于比较。
    空值按 0 处理。
    """
    return normalize_amount_for_excel(value)


def amounts_equal(a: Any, b: Any, tolerance: float = 0.005) -> bool:
    """
    金额近似相等判断。
    """
    return abs(normalize_amount(a) - normalize_amount(b)) <= tolerance


def normalize_date_text(value: Any) -> str:
    """
    将日期统一成 YYYY-MM-DD。
    """
    text = normalize_cell(value)

    if not text:
        return ""

    try:
        return pd.to_datetime(text).strftime("%Y-%m-%d")
    except Exception:
        return text[:10]


def month_end_date(value: Any) -> str:
    """
    将任意日期转为该月月末日期。
    """
    text = normalize_date_text(value)
    if not text:
        return ""

    try:
        dt = pd.to_datetime(text)
        return (dt + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")
    except Exception:
        return ""


def normalize_compare_text(value: Any) -> str:
    """
    用于文本比较的清洗。
    """
    return normalize_cell(value).replace(" ", "").replace("\n", "").replace("\r", "")


def subject_code_text(value: Any) -> str:
    """
    科目编码清洗，避免 400101.0 这类 Excel 读数问题。
    """
    text = normalize_cell(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def safe_int_text(value: Any) -> str:
    """
    年月日字段清洗。
    """
    text = normalize_cell(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def ledger_date_from_row(row: dict[str, Any]) -> str:
    """
    从台账行的 年/月/日 组成 YYYY-MM-DD。
    """
    year = safe_int_text(row.get("年"))
    month = safe_int_text(row.get("月"))
    day = safe_int_text(row.get("日"))

    if not year or not month or not day:
        return ""

    try:
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
    except Exception:
        return ""


def has_ai_review_marker(value: Any) -> bool:
    """
    判断凭证行是否标记为 AI 建议补录 / 待复核。
    """
    text = normalize_cell(value)
    return any(marker in text for marker in AI_REVIEW_MARKERS)


def build_subject_maps(subject_df: pd.DataFrame) -> tuple[set[str], dict[str, str]]:
    """
    从会计科目表提取合法科目编码和科目名称映射。
    """
    valid_codes: set[str] = set()
    name_map: dict[str, str] = {}

    if subject_df is None or subject_df.empty:
        return valid_codes, name_map

    for _, row in subject_df.iterrows():
        code = subject_code_text(row.get("subject_code"))
        name = normalize_cell(row.get("subject_name"))

        if not code:
            continue

        valid_codes.add(code)
        name_map[code] = name

    return valid_codes, name_map


# =========================
# 读取凭证和台账
# =========================

def read_voucher_rows_from_excel(voucher_path: Path) -> list[dict[str, Any]]:
    """
    读取上传的凭证草稿 Excel，并按中文字段名转成行字典。
    """
    wb = load_workbook(voucher_path, data_only=True)
    active_ws = wb.active

    if active_ws is None or not isinstance(active_ws, Worksheet):
        raise ValueError("凭证草稿未找到可用工作表")

    ws = cast(Worksheet, active_ws)

    header_row = find_voucher_header_row(ws)
    data_start_row = header_row + 1
    col_map = get_voucher_column_map(ws, header_row)

    rows: list[dict[str, Any]] = []

    fields = [
        "核算账簿",
        "凭证类别",
        "凭证号",
        "附单据数",
        "制单人",
        "制单日期",
        "审核人",
        "审核日期",
        "摘要",
        "表头自定义项2",
        "表头自定义项3",
        "科目编码",
        "币种",
        "原币借方金额",
        "本币借方金额",
        "原币贷方金额",
        "本币贷方金额",
        "票据号",
        "结算日期",
        "结算方式",
        "核销号",
        "核销业务日期",
    ]

    for row_idx in range(data_start_row, ws.max_row + 1):
        row_dict: dict[str, Any] = {"_excel_row_no": row_idx}
        has_value = False

        for field in fields:
            col_idx = col_map.get(field)
            value = ""

            if col_idx is not None:
                value = ws.cell(row=row_idx, column=col_idx).value

            if normalize_cell(value):
                has_value = True

            row_dict[field] = value

        if has_value:
            rows.append(row_dict)

    return rows


def read_ledger_rows_from_excel(ledger_path: Path) -> list[dict[str, Any]]:
    """
    读取上传的台账草稿 Excel。
    只读取明细区，遇到“合计”或“收入组成”即停止。
    """
    wb = load_workbook(ledger_path, data_only=True)
    active_ws = wb.active

    if active_ws is None or not isinstance(active_ws, Worksheet):
        raise ValueError("台账草稿未找到可用工作表")

    ws = cast(Worksheet, active_ws)

    header_row = find_ledger_header_row(ws)
    data_start_row = header_row + 1
    col_map = get_ledger_column_map(ws, header_row)

    rows: list[dict[str, Any]] = []

    fields = [
        "年",
        "月",
        "日",
        "编号",
        "摘要",
        "支出",
        "收入",
        "余额",
        "标签",
        "科目编码",
        "科目名称",
        "流程编号",
        "对方单位",
        "收支方向",
    ]

    for row_idx in range(data_start_row, ws.max_row + 1):
        row_dict: dict[str, Any] = {"_excel_row_no": row_idx}

        for field in fields:
            col_idx = col_map.get(field)
            value = ""

            if col_idx is not None:
                value = ws.cell(row=row_idx, column=col_idx).value

            row_dict[field] = value

        summary = normalize_compare_text(row_dict.get("摘要"))

        if summary in ["合计", "收入组成：", "收入组成"]:
            break

        has_value = any(
            normalize_cell(row_dict.get(field))
            for field in ["年", "月", "日", "摘要", "支出", "收入", "标签"]
        )

        if has_value:
            rows.append(row_dict)

    return rows


def read_ledger_total_and_income_composition(ledger_path: Path) -> dict[str, Any]:
    """
    读取台账合计行和收入组成区域，供复核检查使用。
    """
    wb = load_workbook(ledger_path, data_only=True)
    active_ws = wb.active

    if active_ws is None or not isinstance(active_ws, Worksheet):
        raise ValueError("台账草稿未找到可用工作表")

    ws = cast(Worksheet, active_ws)

    header_row = find_ledger_header_row(ws)
    col_map = get_ledger_column_map(ws, header_row)

    summary_col = col_map.get("摘要")
    expense_col = col_map.get("支出")
    income_col = col_map.get("收入")

    result = {
        "total_row_no": 0,
        "reported_expense_total": 0.0,
        "reported_income_total": 0.0,
        "income_composition": {},
    }

    if summary_col is None:
        return result

    total_row = 0

    for row_idx in range(header_row + 1, ws.max_row + 1):
        summary = normalize_compare_text(ws.cell(row=row_idx, column=summary_col).value)

        if summary == "合计":
            total_row = row_idx
            result["total_row_no"] = row_idx

            if expense_col is not None:
                result["reported_expense_total"] = normalize_amount(ws.cell(row=row_idx, column=expense_col).value)
            if income_col is not None:
                result["reported_income_total"] = normalize_amount(ws.cell(row=row_idx, column=income_col).value)

            break

    if total_row <= 0:
        return result

    composition: dict[str, float] = {}

    for row_idx in range(total_row + 1, ws.max_row + 1):
        label = normalize_cell(ws.cell(row=row_idx, column=summary_col).value)
        if not label or label == "收入组成：":
            continue

        value = 0.0
        if income_col is not None:
            value = normalize_amount(ws.cell(row=row_idx, column=income_col).value)

        composition[label] = value

    result["income_composition"] = composition
    return result


# =========================
# 凭证业务单元化
# =========================

def _row_debit_amount(row: dict[str, Any]) -> float:
    return normalize_amount(row.get("本币借方金额")) or normalize_amount(row.get("原币借方金额"))


def _row_credit_amount(row: dict[str, Any]) -> float:
    return normalize_amount(row.get("本币贷方金额")) or normalize_amount(row.get("原币贷方金额"))


def _row_subject(row: dict[str, Any]) -> str:
    return subject_code_text(row.get("科目编码"))


def _rows_look_like_pair(row1: dict[str, Any], row2: dict[str, Any]) -> bool:
    """
    判断相邻两行是否构成一笔凭证业务。
    """
    summary1 = normalize_compare_text(row1.get("摘要"))
    summary2 = normalize_compare_text(row2.get("摘要"))

    date1 = normalize_date_text(row1.get("结算日期")) or normalize_date_text(row1.get("制单日期"))
    date2 = normalize_date_text(row2.get("结算日期")) or normalize_date_text(row2.get("制单日期"))

    amount1 = max(_row_debit_amount(row1), _row_credit_amount(row1))
    amount2 = max(_row_debit_amount(row2), _row_credit_amount(row2))

    if not summary1 or not summary2:
        return False

    if summary1 != summary2:
        return False

    if date1 != date2:
        return False

    if not amounts_equal(amount1, amount2):
        return False

    return True


def build_voucher_unit_from_rows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """
    将一笔凭证业务的借贷行合并成业务单元。
    """
    debit_total = sum(_row_debit_amount(row) for row in rows)
    credit_total = sum(_row_credit_amount(row) for row in rows)

    debit_subjects = [_row_subject(row) for row in rows if _row_debit_amount(row) > 0]
    credit_subjects = [_row_subject(row) for row in rows if _row_credit_amount(row) > 0]

    non_bank_subject = ""
    direction = "未知"

    if BANK_SUBJECT_CODE in debit_subjects and any(code != BANK_SUBJECT_CODE for code in credit_subjects):
        direction = "收入"
        non_bank_subject = next((code for code in credit_subjects if code != BANK_SUBJECT_CODE), "")
    elif BANK_SUBJECT_CODE in credit_subjects and any(code != BANK_SUBJECT_CODE for code in debit_subjects):
        direction = "支出"
        non_bank_subject = next((code for code in debit_subjects if code != BANK_SUBJECT_CODE), "")

    first = rows[0]
    amount = debit_total if debit_total > 0 else credit_total

    return {
        "row_nos": [row.get("_excel_row_no") for row in rows],
        "summary": normalize_cell(first.get("摘要")),
        "summary_key": normalize_compare_text(first.get("摘要")),
        "direction": direction,
        "amount": round(amount, 2),
        "debit_total": round(debit_total, 2),
        "credit_total": round(credit_total, 2),
        "debit_subjects": debit_subjects,
        "credit_subjects": credit_subjects,
        "subject_code": non_bank_subject,
        "voucher_date": normalize_date_text(first.get("制单日期")),
        "settlement_date": normalize_date_text(first.get("结算日期")),
        "writeoff_date": normalize_date_text(first.get("核销业务日期")),
        "custom_no": normalize_cell(first.get("表头自定义项3")),
        "ai_review_flag": any(has_ai_review_marker(row.get("表头自定义项2")) for row in rows),
    }


def build_voucher_business_units(actual_rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """
    将上传凭证行转换成业务单元。
    不再与系统标准凭证逐行对比，避免 AI 补录行导致整体错位。
    """
    units: list[dict[str, Any]] = []
    exceptions: list[dict[str, Any]] = []

    idx = 0

    while idx < len(actual_rows):
        current = actual_rows[idx]

        if idx + 1 < len(actual_rows) and _rows_look_like_pair(current, actual_rows[idx + 1]):
            unit_rows = [current, actual_rows[idx + 1]]
            idx += 2
        else:
            unit_rows = [current]
            idx += 1

        unit = build_voucher_unit_from_rows(unit_rows)
        units.append(unit)

        if len(unit_rows) != 2:
            exceptions.append({
                "type": "VOUCHER_PAIR_STRUCTURE_WARNING",
                "message": f"凭证第 {unit['row_nos']} 行未能识别为标准借贷两行结构。",
                "suggestion": "请检查该业务是否缺少借方或贷方行；如果是特殊分录，请人工复核。"
            })

        if not amounts_equal(unit.get("debit_total"), unit.get("credit_total")):
            exceptions.append({
                "type": "VOUCHER_BUSINESS_NOT_BALANCED",
                "message": f"凭证业务行 {unit['row_nos']} 借贷不平：借方 {unit.get('debit_total')}，贷方 {unit.get('credit_total')}。",
                "suggestion": "请检查该笔业务借贷金额是否一致。"
            })

        if unit.get("direction") not in ["收入", "支出"]:
            exceptions.append({
                "type": "VOUCHER_DIRECTION_UNCLEAR",
                "message": f"凭证业务行 {unit['row_nos']} 无法识别借贷方向，应为银行存款1002与一个非银行科目对应。",
                "suggestion": "收入业务应借记1002、贷记收入科目；支出业务应借记支出科目、贷记1002。"
            })

    return units, exceptions


# =========================
# 流水与凭证业务匹配
# =========================

def flow_to_dict(row: pd.Series) -> dict[str, Any]:
    return {str(key): value for key, value in row.to_dict().items()}


def score_unit_for_flow(flow: dict[str, Any], unit: dict[str, Any]) -> int:
    """
    在方向、金额、日期已匹配的基础上，用关键词进一步评分。
    """
    score = 0

    summary = normalize_compare_text(unit.get("summary"))
    counterparty = normalize_compare_text(flow.get("counterparty"))
    oa_flow_no = normalize_compare_text(flow.get("oa_flow_no"))
    combined_text = normalize_compare_text(flow.get("combined_text"))

    if counterparty and counterparty in summary:
        score += 5

    if oa_flow_no and oa_flow_no in summary:
        score += 5

    if "利息" in combined_text and "利息" in summary:
        score += 4

    if "党费" in combined_text and "党费" in summary:
        score += 3

    return score


def match_flows_to_voucher_units(
    bank_df: pd.DataFrame,
    voucher_units: list[dict[str, Any]],
    original_exception_flow_indices: set[str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """
    用业务实质匹配银行流水与上传凭证业务单元：
    - 方向一致
    - 金额一致
    - 结算日期 = 流水交易日期
    不再要求与“系统标准结果”逐行一致。
    """
    exceptions: list[dict[str, Any]] = []
    ai_supplements: list[dict[str, Any]] = []
    matched_links: list[dict[str, Any]] = []

    used_unit_indices: set[int] = set()

    for _, flow_row in bank_df.iterrows():
        flow = flow_to_dict(flow_row)

        flow_index = normalize_cell(flow.get("flow_index"))
        flow_date = normalize_date_text(flow.get("transaction_date"))
        direction = normalize_cell(flow.get("direction"))
        amount = normalize_amount(flow.get("amount"))

        candidate_indices: list[tuple[int, int]] = []

        for unit_idx, unit in enumerate(voucher_units):
            if unit_idx in used_unit_indices:
                continue

            if unit.get("direction") != direction:
                continue

            if not amounts_equal(unit.get("amount"), amount):
                continue

            if normalize_date_text(unit.get("settlement_date")) != flow_date:
                continue

            score = score_unit_for_flow(flow, unit)

            # 原始异常业务优先匹配带 AI 待复核标记的凭证单元
            if flow_index in original_exception_flow_indices and unit.get("ai_review_flag"):
                score += 10

            candidate_indices.append((score, unit_idx))

        if not candidate_indices:
            exceptions.append({
                "type": "FLOW_NOT_COVERED_BY_VOUCHER",
                "message": f"流水序号 {flow_index} 未在上传凭证中找到对应业务：日期 {flow_date}，方向 {direction}，金额 {amount}，对方单位/户名：{normalize_cell(flow.get('counterparty'))}。",
                "suggestion": "请检查该流水是否漏生成凭证；若属于AI补录业务，应在凭证中有对应借贷两行并标注“AI建议补录-待复核”。"
            })
            continue

        candidate_indices.sort(key=lambda x: x[0], reverse=True)
        _, selected_idx = candidate_indices[0]
        used_unit_indices.add(selected_idx)

        unit = voucher_units[selected_idx]
        unit["matched_flow_index"] = flow_index
        unit["matched_flow"] = flow

        matched_links.append({
            "flow_index": flow_index,
            "unit_index": selected_idx,
            "flow": flow,
            "unit": unit,
        })

        if flow_index in original_exception_flow_indices or unit.get("ai_review_flag"):
            if not unit.get("ai_review_flag"):
                exceptions.append({
                    "type": "AI_SUPPLEMENT_MARK_MISSING",
                    "message": f"流水序号 {flow_index} 属于原始规则未能自动匹配的业务，但凭证对应分录未标注“AI建议补录-待复核”。",
                    "suggestion": "请在AI补录分录的表头自定义项2中保留“AI建议补录-待复核”，便于人工复核。"
                })

            ai_supplements.append({
                "flow_index": flow_index,
                "transaction_date": flow_date,
                "direction": direction,
                "counterparty": normalize_cell(flow.get("counterparty")),
                "amount": amount,
                "subject_code": unit.get("subject_code", ""),
                "summary": unit.get("summary", ""),
                "voucher_rows": unit.get("row_nos", []),
                "review_prompt": "该业务由AI建议补录或对应原始异常流水，已在本次复核中按业务实质匹配通过，但仍需财务人员结合银行回单、OA流程、党费收缴明细或业务说明人工确认后方可正式入账。",
            })

    # 上传凭证中存在无法对应任何流水的业务
    for unit_idx, unit in enumerate(voucher_units):
        if unit_idx in used_unit_indices:
            continue

        # 空白或无法识别的结构异常前面已经提示，这里避免重复过度提示
        exceptions.append({
            "type": "VOUCHER_UNIT_NOT_MATCHED_TO_FLOW",
            "message": f"凭证业务行 {unit.get('row_nos')} 未匹配到对应银行流水：摘要“{unit.get('summary')}”，方向 {unit.get('direction')}，金额 {unit.get('amount')}，结算日期 {unit.get('settlement_date')}。",
            "suggestion": "请检查是否存在重复补录、错误凭证行，或结算日期/金额与银行流水不一致。"
        })

    return matched_links, ai_supplements, exceptions


# =========================
# 凭证科目与日期检查
# =========================

def validate_voucher_units(
    voucher_units: list[dict[str, Any]],
    valid_subject_codes: set[str],
) -> list[dict[str, Any]]:
    """
    检查凭证业务单元的基本合法性。
    """
    exceptions: list[dict[str, Any]] = []

    for unit in voucher_units:
        subject_code = subject_code_text(unit.get("subject_code"))

        if subject_code and valid_subject_codes and subject_code not in valid_subject_codes:
            exceptions.append({
                "type": "VOUCHER_SUBJECT_NOT_IN_SUBJECT_TABLE",
                "message": f"凭证业务行 {unit.get('row_nos')} 使用的非银行科目编码 {subject_code} 不存在于会计科目表。",
                "suggestion": "请检查AI建议或业务映射规则是否使用了会计科目表以外的科目。"
            })

        settlement_date = normalize_date_text(unit.get("settlement_date"))
        writeoff_date = normalize_date_text(unit.get("writeoff_date"))

        if settlement_date and writeoff_date and settlement_date != writeoff_date:
            exceptions.append({
                "type": "VOUCHER_DATE_INCONSISTENT",
                "message": f"凭证业务行 {unit.get('row_nos')} 结算日期 {settlement_date} 与核销业务日期 {writeoff_date} 不一致。",
                "suggestion": "通常结算日期和核销业务日期应等于银行流水交易日期。"
            })

    return exceptions


def check_voucher_balance(actual_rows: list[dict[str, Any]]) -> dict[str, Any]:
    """
    检查上传凭证整体借贷是否平衡。
    """
    debit_total = 0.0
    credit_total = 0.0

    for row in actual_rows:
        debit_total += normalize_amount(row.get("本币借方金额"))
        credit_total += normalize_amount(row.get("本币贷方金额"))

    balance_check = abs(debit_total - credit_total) <= 0.005

    return {
        "debit_total": round(debit_total, 2),
        "credit_total": round(credit_total, 2),
        "balance_check": balance_check,
    }


# =========================
# 台账检查
# =========================

def ledger_row_direction(row: dict[str, Any]) -> str:
    income = normalize_amount(row.get("收入"))
    expense = normalize_amount(row.get("支出"))

    if income > 0 and expense > 0:
        return "双向"
    if income > 0:
        return "收入"
    if expense > 0:
        return "支出"
    return "未知"


def ledger_row_amount(row: dict[str, Any]) -> float:
    return max(normalize_amount(row.get("收入")), normalize_amount(row.get("支出")))


def validate_ledger_rows_are_single_direction(actual_ledger_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    检查台账每一行只能填收入或支出其中一列。
    """
    exceptions: list[dict[str, Any]] = []

    for row in actual_ledger_rows:
        income = normalize_amount(row.get("收入"))
        expense = normalize_amount(row.get("支出"))

        if income > 0 and expense > 0:
            exceptions.append({
                "type": "LEDGER_ROW_HAS_BOTH_INCOME_AND_EXPENSE",
                "message": f"台账第 {row.get('_excel_row_no')} 行同时存在支出 {expense} 和收入 {income}：摘要“{normalize_cell(row.get('摘要'))}”。",
                "suggestion": "每一条台账明细只能对应一个方向：收入业务只填收入列，支出业务只填支出列；请清空错误方向的金额。"
            })

    return exceptions


def match_voucher_units_to_ledger_rows(
    voucher_units: list[dict[str, Any]],
    actual_ledger_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """
    检查每个凭证业务单元是否有对应台账行。
    """
    exceptions: list[dict[str, Any]] = []
    matched_ledger_links: list[dict[str, Any]] = []
    used_ledger_indices: set[int] = set()

    for unit_idx, unit in enumerate(voucher_units):
        direction = normalize_cell(unit.get("direction"))
        amount = normalize_amount(unit.get("amount"))
        voucher_date = normalize_date_text(unit.get("voucher_date"))
        summary_key = normalize_compare_text(unit.get("summary"))

        candidate_indices: list[tuple[int, int]] = []

        for ledger_idx, ledger_row in enumerate(actual_ledger_rows):
            if ledger_idx in used_ledger_indices:
                continue

            row_direction = ledger_row_direction(ledger_row)
            if row_direction != direction:
                continue

            if not amounts_equal(ledger_row_amount(ledger_row), amount):
                continue

            if ledger_date_from_row(ledger_row) != voucher_date:
                continue

            ledger_summary_key = normalize_compare_text(ledger_row.get("摘要"))
            score = 0
            if ledger_summary_key == summary_key:
                score += 10
            elif summary_key and (summary_key in ledger_summary_key or ledger_summary_key in summary_key):
                score += 5

            candidate_indices.append((score, ledger_idx))

        if not candidate_indices:
            exceptions.append({
                "type": "LEDGER_ROW_NOT_FOUND_FOR_VOUCHER",
                "message": f"凭证业务行 {unit.get('row_nos')} 未在台账明细中找到对应记录：摘要“{unit.get('summary')}”，方向 {direction}，金额 {amount}，台账日期应为 {voucher_date}。",
                "suggestion": "请检查台账是否漏记、收入/支出列是否填反，或台账年月日是否与凭证制单日期一致。"
            })
            continue

        candidate_indices.sort(key=lambda x: x[0], reverse=True)
        _, selected_idx = candidate_indices[0]
        used_ledger_indices.add(selected_idx)

        matched_ledger_links.append({
            "unit_index": unit_idx,
            "ledger_index": selected_idx,
            "unit": unit,
            "ledger_row": actual_ledger_rows[selected_idx],
        })

    for ledger_idx, ledger_row in enumerate(actual_ledger_rows):
        if ledger_idx in used_ledger_indices:
            continue

        direction = ledger_row_direction(ledger_row)
        if direction == "未知":
            continue

        exceptions.append({
            "type": "LEDGER_ROW_NOT_MATCHED_TO_VOUCHER",
            "message": f"台账第 {ledger_row.get('_excel_row_no')} 行未匹配到对应凭证业务：摘要“{normalize_cell(ledger_row.get('摘要'))}”，方向 {direction}，金额 {ledger_row_amount(ledger_row)}。",
            "suggestion": "请检查台账是否存在重复行、错误金额，或凭证是否缺少对应分录。"
        })

    return matched_ledger_links, exceptions


def validate_ledger_totals_and_income_composition(
    ledger_path: Path,
    actual_ledger_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    检查台账合计行和收入组成是否与明细一致。
    """
    exceptions: list[dict[str, Any]] = []

    total_info = read_ledger_total_and_income_composition(ledger_path)

    calculated_income_total = round(sum(normalize_amount(row.get("收入")) for row in actual_ledger_rows), 2)
    calculated_expense_total = round(sum(normalize_amount(row.get("支出")) for row in actual_ledger_rows), 2)

    reported_income_total = total_info.get("reported_income_total", 0.0)
    reported_expense_total = total_info.get("reported_expense_total", 0.0)

    if not amounts_equal(calculated_income_total, reported_income_total):
        exceptions.append({
            "type": "LEDGER_INCOME_TOTAL_MISMATCH",
            "message": f"台账收入合计不一致：明细收入合计为 {calculated_income_total}，合计行收入为 {reported_income_total}。",
            "suggestion": "请重新计算台账收入合计，确保AI补录收入也纳入合计。"
        })

    if not amounts_equal(calculated_expense_total, reported_expense_total):
        exceptions.append({
            "type": "LEDGER_EXPENSE_TOTAL_MISMATCH",
            "message": f"台账支出合计不一致：明细支出合计为 {calculated_expense_total}，合计行支出为 {reported_expense_total}。",
            "suggestion": "请重新计算台账支出合计。"
        })

    income_by_tag: dict[str, float] = {}
    for row in actual_ledger_rows:
        tag = normalize_cell(row.get("标签"))
        income = normalize_amount(row.get("收入"))
        if income > 0 and tag:
            income_by_tag[tag] = round(income_by_tag.get(tag, 0.0) + income, 2)

    composition = total_info.get("income_composition", {}) or {}

    for tag, calculated_value in income_by_tag.items():
        reported_value = normalize_amount(composition.get(tag))
        if not amounts_equal(calculated_value, reported_value):
            exceptions.append({
                "type": "LEDGER_INCOME_COMPOSITION_MISMATCH",
                "message": f"台账收入组成【{tag}】不一致：明细汇总为 {calculated_value}，收入组成区域为 {reported_value}。",
                "suggestion": "请重新计算收入组成，确保AI补录收入按标签纳入统计。"
            })

    return exceptions


# =========================
# 报告文本
# =========================

def build_exception_summary(
    review_exceptions: list[dict[str, Any]],
    max_items: int = 5
) -> str:
    """
    生成异常摘要文本。
    """
    if not review_exceptions:
        return ""

    summary_lines = []

    for idx, item in enumerate(review_exceptions[:max_items], start=1):
        message = normalize_cell(item.get("message"))
        suggestion = normalize_cell(item.get("suggestion"))

        if suggestion:
            summary_lines.append(f"{idx}）{message}建议：{suggestion}")
        else:
            summary_lines.append(f"{idx}）{message}")

    remaining_count = len(review_exceptions) - max_items

    if remaining_count > 0:
        summary_lines.append(f"其余 {remaining_count} 项异常详见“复核异常清单”工作表。")

    return "主要异常包括：" + "；".join(summary_lines)


def build_ai_supplement_text(ai_supplement_items: list[dict[str, Any]]) -> str:
    """
    生成 AI 补录提示文本。
    """
    if not ai_supplement_items:
        return "本次未识别到AI建议补录业务。"

    lines = []
    for idx, item in enumerate(ai_supplement_items, start=1):
        lines.append(
            f"{idx}. 流水序号{item.get('flow_index')}，日期{item.get('transaction_date')}，"
            f"方向{item.get('direction')}，对方单位/户名：{item.get('counterparty')}，"
            f"金额：{item.get('amount')}，凭证科目：{item.get('subject_code')}，"
            f"摘要：{item.get('summary')}。复核提示：{item.get('review_prompt')}"
        )

    return "\n".join(lines)


def build_review_report(
    review_status: str,
    matched_count: int,
    original_exception_count: int,
    voucher_balance: dict[str, Any],
    review_exceptions: list[dict[str, Any]],
    ai_supplement_items: list[dict[str, Any]],
    covered_flow_count: int,
    voucher_business_count: int,
    ledger_detail_count: int,
) -> str:
    """
    生成复核报告文本。
    新版逻辑采用“业务实质复核”，不再和基础系统标准结果逐行比较。
    """
    ai_count = len(ai_supplement_items)

    if review_status == "passed":
        return (
            f"系统复核无异常。本次采用业务实质复核逻辑，根据银行流水、OA流程、党费业务映射规则表、"
            f"党员离退休情况表和会计科目表，对上传凭证与台账进行核对。"
            f"本次银行流水覆盖 {covered_flow_count} 笔，上传凭证识别业务 {voucher_business_count} 笔，"
            f"上传台账识别明细 {ledger_detail_count} 笔；"
            f"其中基础规则成功匹配业务 {matched_count} 笔，原始业务异常 {original_exception_count} 笔，"
            f"识别到 AI 建议补录业务 {ai_count} 笔。"
            f"凭证借方合计 {voucher_balance['debit_total']}，贷方合计 {voucher_balance['credit_total']}，"
            f"借贷平衡校验通过。AI补录业务不计为复核异常，但必须由财务人员人工确认后方可正式入账。"
        )

    exception_summary = build_exception_summary(review_exceptions, max_items=5)

    return (
        f"系统复核发现异常。本次采用业务实质复核逻辑，不再要求上传凭证/台账与基础规则生成结果逐行一致，"
        f"而是核对银行流水、OA流程、规则表、党员状态表、会计科目表与上传凭证/台账之间的业务覆盖关系。"
        f"本次发现 {len(review_exceptions)} 项复核异常；银行流水覆盖 {covered_flow_count} 笔，"
        f"上传凭证识别业务 {voucher_business_count} 笔，上传台账识别明细 {ledger_detail_count} 笔；"
        f"基础规则成功匹配业务 {matched_count} 笔，原始业务异常 {original_exception_count} 笔，"
        f"识别到 AI 建议补录业务 {ai_count} 笔。"
        f"凭证借方合计 {voucher_balance['debit_total']}，贷方合计 {voucher_balance['credit_total']}，"
        f"借贷平衡校验结果为{'通过' if voucher_balance['balance_check'] else '不通过'}。"
        f"{exception_summary}"
        f"AI补录业务本身不作为异常，但需财务人员结合银行回单、OA流程、党费收缴明细或业务说明人工确认。"
    )


# =========================
# 主复核入口
# =========================

def perform_review(
    bank_df: pd.DataFrame,
    oa_df: pd.DataFrame,
    subject_df: pd.DataFrame,
    member_df: pd.DataFrame,
    rule_df: pd.DataFrame,
    voucher_path: Path,
    ledger_path: Path,
    maker: str = "",
    book_code: str = "",
    voucher_type: str = "",
) -> dict[str, Any]:
    """
    执行凭证和台账复核。

    新版复核逻辑：
    1. 不再用“系统标准结果”逐行比对上传凭证/台账；
    2. 改为根据银行流水、OA流程、规则表、党员状态表、会计科目表进行业务实质核对；
    3. AI建议补录业务只做待人工复核提示，不列为异常；
    4. 只有真正的借贷不平、金额/日期/方向不一致、台账漏记、科目不存在等问题才列为异常。
    """
    matched_df, original_business_exceptions = match_all_business(
        bank_df=bank_df,
        oa_df=oa_df,
        subject_df=subject_df,
        member_df=member_df,
        rule_df=rule_df
    )

    matched_count = int((matched_df["match_status"] == "matched").sum())
    original_exception_count = int(len(original_business_exceptions))
    original_exception_flow_indices = {
        normalize_cell(item.get("flow_index"))
        for item in original_business_exceptions
    }

    valid_subject_codes, _subject_name_map = build_subject_maps(subject_df)

    actual_voucher_rows = read_voucher_rows_from_excel(voucher_path)
    actual_ledger_rows = read_ledger_rows_from_excel(ledger_path)

    voucher_balance = check_voucher_balance(actual_voucher_rows)

    review_exceptions: list[dict[str, Any]] = []

    if not voucher_balance["balance_check"]:
        review_exceptions.append({
            "type": "VOUCHER_NOT_BALANCED",
            "message": f"上传凭证借贷不平：借方合计 {voucher_balance['debit_total']}，贷方合计 {voucher_balance['credit_total']}。",
            "suggestion": "请检查凭证金额是否漏填、错填，或是否误删凭证明细行。"
        })

    voucher_units, voucher_structure_exceptions = build_voucher_business_units(actual_voucher_rows)
    review_exceptions.extend(voucher_structure_exceptions)
    review_exceptions.extend(validate_voucher_units(voucher_units, valid_subject_codes))

    matched_links, ai_supplement_items, flow_voucher_exceptions = match_flows_to_voucher_units(
        bank_df=bank_df,
        voucher_units=voucher_units,
        original_exception_flow_indices=original_exception_flow_indices,
    )
    review_exceptions.extend(flow_voucher_exceptions)

    review_exceptions.extend(validate_ledger_rows_are_single_direction(actual_ledger_rows))

    _ledger_links, ledger_match_exceptions = match_voucher_units_to_ledger_rows(
        voucher_units=voucher_units,
        actual_ledger_rows=actual_ledger_rows,
    )
    review_exceptions.extend(ledger_match_exceptions)

    review_exceptions.extend(
        validate_ledger_totals_and_income_composition(
            ledger_path=ledger_path,
            actual_ledger_rows=actual_ledger_rows,
        )
    )

    review_status = "passed" if not review_exceptions else "failed"

    review_report = build_review_report(
        review_status=review_status,
        matched_count=matched_count,
        original_exception_count=original_exception_count,
        voucher_balance=voucher_balance,
        review_exceptions=review_exceptions,
        ai_supplement_items=ai_supplement_items,
        covered_flow_count=len(matched_links),
        voucher_business_count=len(voucher_units),
        ledger_detail_count=len(actual_ledger_rows),
    )

    return {
        "review_status": review_status,
        "review_passed": review_status == "passed",
        "review_report": review_report,
        "review_exception_count": len(review_exceptions),
        "review_exceptions": review_exceptions,
        "original_business_exception_count": original_exception_count,
        "original_business_exceptions": original_business_exceptions,
        "matched_count": matched_count,

        # 为兼容 app.py / Dify 原解析逻辑，保留这些字段。
        # 新版含义不再是“系统标准行数”，而是“上传文件业务实质识别结果”。
        "expected_voucher_row_count": len(actual_voucher_rows),
        "actual_voucher_row_count": len(actual_voucher_rows),
        "expected_ledger_row_count": len(actual_ledger_rows),
        "actual_ledger_row_count": len(actual_ledger_rows),

        "voucher_balance": voucher_balance,

        # 新增：AI补录提示和业务实质核对统计
        "ai_supplement_count": len(ai_supplement_items),
        "ai_supplement_items": ai_supplement_items,
        "ai_supplement_text": build_ai_supplement_text(ai_supplement_items),
        "covered_flow_count": len(matched_links),
        "voucher_business_count": len(voucher_units),
        "ledger_detail_count": len(actual_ledger_rows),
    }
