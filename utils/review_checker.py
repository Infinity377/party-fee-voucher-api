from __future__ import annotations

from pathlib import Path
from typing import Any, cast
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from utils.cleaners import normalize_cell
from utils.matcher import match_all_business
from utils.voucher_generator import (
    build_voucher_rows,
    find_header_row as find_voucher_header_row,
    get_template_column_map as get_voucher_column_map,
    normalize_amount_for_excel,
)
from utils.ledger_generator import (
    build_ledger_rows,
    find_ledger_header_row,
    get_ledger_column_map,
)


def normalize_amount(value: Any) -> float:
    """
    统一金额格式，便于比较。
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


def normalize_compare_text(value: Any) -> str:
    """
    用于文本比较的清洗。
    """
    return normalize_cell(value).replace(" ", "").replace("\n", "").replace("\r", "")


def read_voucher_rows_from_excel(voucher_path: Path) -> list[dict[str, Any]]:
    """
    读取上传的凭证草稿 Excel，并按中文字段名转成行字典。
    """
    wb = load_workbook(voucher_path, data_only=True)
    ws = cast(Worksheet, wb.active)

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
        row_dict: dict[str, Any] = {}

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
    ws = cast(Worksheet, wb.active)

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
        row_dict: dict[str, Any] = {}

        for field in fields:
            col_idx = col_map.get(field)
            value = ""

            if col_idx is not None:
                value = ws.cell(row=row_idx, column=col_idx).value

            row_dict[field] = value

        summary = normalize_compare_text(row_dict.get("摘要"))

        if summary in ["合计", "收入组成：", "收入组成"]:
            break

        has_value = any(normalize_cell(row_dict.get(field)) for field in ["年", "月", "日", "摘要", "支出", "收入", "标签"])

        if has_value:
            rows.append(row_dict)

    return rows


def compare_voucher_rows(
    expected_rows: list[dict[str, Any]],
    actual_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    比对凭证明细。
    """
    exceptions: list[dict[str, Any]] = []

    if len(expected_rows) != len(actual_rows):
        exceptions.append(
            {
                "type": "VOUCHER_ROW_COUNT_MISMATCH",
                "message": f"凭证行数不一致：系统应为 {len(expected_rows)} 行，上传凭证为 {len(actual_rows)} 行。",
                "suggestion": "请检查是否误删、误增凭证明细行，或是否上传了错误版本的凭证草稿。"
            }
        )

    compare_len = min(len(expected_rows), len(actual_rows))

    text_fields = [
        "核算账簿",
        "凭证类别",
        "制单人",
        "摘要",
        "表头自定义项3",
        "科目编码",
        "币种",
    ]

    date_fields = [
        "制单日期",
        "结算日期",
        "核销业务日期",
    ]

    amount_fields = [
        "原币借方金额",
        "本币借方金额",
        "原币贷方金额",
        "本币贷方金额",
    ]

    for idx in range(compare_len):
        expected = expected_rows[idx]
        actual = actual_rows[idx]
        row_no = idx + 1

        for field in text_fields:
            expected_value = normalize_compare_text(expected.get(field))
            actual_value = normalize_compare_text(actual.get(field))

            if expected_value != actual_value:
                exceptions.append(
                    {
                        "type": "VOUCHER_FIELD_MISMATCH",
                        "message": f"凭证第 {row_no} 行字段【{field}】不一致：应为“{expected_value}”，实际为“{actual_value}”。",
                        "suggestion": "请检查凭证草稿是否被人工修改，或业务映射规则是否已更新。"
                    }
                )

        for field in date_fields:
            expected_value = normalize_date_text(expected.get(field))
            actual_value = normalize_date_text(actual.get(field))

            if expected_value != actual_value:
                exceptions.append(
                    {
                        "type": "VOUCHER_DATE_MISMATCH",
                        "message": f"凭证第 {row_no} 行日期字段【{field}】不一致：应为“{expected_value}”，实际为“{actual_value}”。",
                        "suggestion": "请检查制单日期、结算日期、核销业务日期是否按规则填写。"
                    }
                )

        for field in amount_fields:
            if not amounts_equal(expected.get(field), actual.get(field)):
                exceptions.append(
                    {
                        "type": "VOUCHER_AMOUNT_MISMATCH",
                        "message": f"凭证第 {row_no} 行金额字段【{field}】不一致：应为“{expected.get(field)}”，实际为“{actual.get(field)}”。",
                        "suggestion": "请检查凭证金额是否与银行流水金额一致。"
                    }
                )

    return exceptions


def check_voucher_balance(actual_rows: list[dict[str, Any]]) -> dict[str, Any]:
    """
    检查上传凭证借贷是否平衡。
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


def compare_ledger_rows(
    expected_rows: list[dict[str, Any]],
    actual_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    比对台账明细。
    """
    exceptions: list[dict[str, Any]] = []

    if len(expected_rows) != len(actual_rows):
        exceptions.append(
            {
                "type": "LEDGER_ROW_COUNT_MISMATCH",
                "message": f"台账明细行数不一致：系统应为 {len(expected_rows)} 行，上传台账为 {len(actual_rows)} 行。",
                "suggestion": "请检查是否误删、误增台账明细行，或是否上传了错误版本的台账草稿。"
            }
        )

    compare_len = min(len(expected_rows), len(actual_rows))

    text_fields = [
        "年",
        "月",
        "日",
        "摘要",
        "标签",
    ]

    amount_fields = [
        "收入",
        "支出",
    ]

    for idx in range(compare_len):
        expected = expected_rows[idx]
        actual = actual_rows[idx]
        row_no = idx + 1

        for field in text_fields:
            expected_value = normalize_compare_text(expected.get(field))
            actual_value = normalize_compare_text(actual.get(field))

            if expected_value != actual_value:
                exceptions.append(
                    {
                        "type": "LEDGER_FIELD_MISMATCH",
                        "message": f"台账第 {row_no} 行字段【{field}】不一致：应为“{expected_value}”，实际为“{actual_value}”。",
                        "suggestion": "请检查台账日期、摘要或标签是否与凭证生成结果一致。"
                    }
                )

        for field in amount_fields:
            if not amounts_equal(expected.get(field), actual.get(field)):
                exceptions.append(
                    {
                        "type": "LEDGER_AMOUNT_MISMATCH",
                        "message": f"台账第 {row_no} 行金额字段【{field}】不一致：应为“{expected.get(field)}”，实际为“{actual.get(field)}”。",
                        "suggestion": "请检查台账收入/支出金额是否与凭证及流水一致。"
                    }
                )

    return exceptions
def build_exception_summary(
    review_exceptions: list[dict[str, Any]],
    max_items: int = 3
) -> str:
    """
    生成异常摘要文本。
    默认提取前 3 条异常，放进复核报告总述里。
    详细异常仍以“复核异常清单”sheet 为准。
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

def build_review_report(
    review_status: str,
    matched_count: int,
    original_exception_count: int,
    voucher_balance: dict[str, Any],
    review_exceptions: list[dict[str, Any]],
) -> str:
    """
    生成复核报告文本。
    """
    if review_status == "passed":
        return (
            f"系统复核无异常。系统根据银行流水、OA流程和配置规则重新计算后，"
            f"确认上传的凭证草稿和台账草稿与系统标准结果一致。"
            f"本次成功匹配业务 {matched_count} 笔，原始异常业务 {original_exception_count} 笔；"
            f"凭证借方合计 {voucher_balance['debit_total']}，贷方合计 {voucher_balance['credit_total']}，"
            f"借贷平衡校验通过。仍建议由财务人员进行最终人工确认。"
        )

    exception_summary = build_exception_summary(review_exceptions, max_items=3)

    return (
        f"系统复核发现异常。系统根据银行流水、OA流程和配置规则重新计算后，"
        f"发现上传的凭证草稿或台账草稿存在 {len(review_exceptions)} 项不一致。"
        f"本次成功匹配业务 {matched_count} 笔，原始异常业务 {original_exception_count} 笔；"
        f"凭证借方合计 {voucher_balance['debit_total']}，贷方合计 {voucher_balance['credit_total']}，"
        f"借贷平衡校验结果为{'通过' if voucher_balance['balance_check'] else '不通过'}。"
        f"{exception_summary}"
        f"详细信息请查看复核报告 Excel 中的“复核异常清单”工作表，并根据异常清单逐项核对。"
    )

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

    expected_voucher_rows = build_voucher_rows(
        matched_df=matched_df,
        maker=maker,
        book_code=book_code,
        voucher_type=voucher_type
    )

    expected_ledger_rows = build_ledger_rows(matched_df)

    actual_voucher_rows = read_voucher_rows_from_excel(voucher_path)
    actual_ledger_rows = read_ledger_rows_from_excel(ledger_path)

    voucher_balance = check_voucher_balance(actual_voucher_rows)

    review_exceptions: list[dict[str, Any]] = []

    if not voucher_balance["balance_check"]:
        review_exceptions.append(
            {
                "type": "VOUCHER_NOT_BALANCED",
                "message": f"上传凭证借贷不平：借方合计 {voucher_balance['debit_total']}，贷方合计 {voucher_balance['credit_total']}。",
                "suggestion": "请检查凭证金额是否漏填、错填，或是否误删凭证明细行。"
            }
        )

    review_exceptions.extend(
        compare_voucher_rows(
            expected_rows=expected_voucher_rows,
            actual_rows=actual_voucher_rows
        )
    )

    review_exceptions.extend(
        compare_ledger_rows(
            expected_rows=expected_ledger_rows,
            actual_rows=actual_ledger_rows
        )
    )

    review_status = "passed" if not review_exceptions else "failed"

    review_report = build_review_report(
        review_status=review_status,
        matched_count=matched_count,
        original_exception_count=original_exception_count,
        voucher_balance=voucher_balance,
        review_exceptions=review_exceptions
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
        "expected_voucher_row_count": len(expected_voucher_rows),
        "actual_voucher_row_count": len(actual_voucher_rows),
        "expected_ledger_row_count": len(expected_ledger_rows),
        "actual_ledger_row_count": len(actual_ledger_rows),
        "voucher_balance": voucher_balance,
    }