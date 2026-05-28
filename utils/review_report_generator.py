from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.worksheet import Worksheet


def normalize_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    将异常记录转为适合写入 Excel 的普通字典。
    """
    normalized: list[dict[str, Any]] = []

    for item in records:
        normalized.append({
            str(key): "" if value is None else value
            for key, value in item.items()
        })

    return normalized


def autosize_columns(ws: Worksheet) -> None:
    """
    自动调整列宽。
    使用列序号生成列字母，避免 MergedCell 没有 column_letter 属性导致 Pylance 报错。
    """
    from openpyxl.utils import get_column_letter

    for col_idx in range(1, ws.max_column + 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)

        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue

            max_length = max(max_length, len(str(value)))

        adjusted_width = min(max(max_length + 2, 12), 60)
        ws.column_dimensions[column_letter].width = adjusted_width


def style_worksheet(ws: Worksheet) -> None:
    """
    简单美化工作表。
    """
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)
    thin_side = Side(style="thin", color="D9D9D9")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = border

    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    autosize_columns(ws)


def build_summary_rows(review_result: dict[str, Any]) -> list[dict[str, Any]]:
    """
    生成复核结论工作表数据。
    """
    voucher_balance = review_result.get("voucher_balance", {}) or {}

    return [
        {"项目": "复核状态", "结果": review_result.get("review_status", "")},
        {"项目": "是否通过", "结果": "通过" if review_result.get("review_passed") else "不通过"},
        {"项目": "复核异常数量", "结果": review_result.get("review_exception_count", 0)},
        {"项目": "成功匹配业务数量", "结果": review_result.get("matched_count", 0)},
        {"项目": "原始业务异常数量", "结果": review_result.get("original_business_exception_count", 0)},
        {"项目": "理论凭证行数", "结果": review_result.get("expected_voucher_row_count", 0)},
        {"项目": "上传凭证行数", "结果": review_result.get("actual_voucher_row_count", 0)},
        {"项目": "理论台账明细行数", "结果": review_result.get("expected_ledger_row_count", 0)},
        {"项目": "上传台账明细行数", "结果": review_result.get("actual_ledger_row_count", 0)},
        {"项目": "凭证借方合计", "结果": voucher_balance.get("debit_total", 0)},
        {"项目": "凭证贷方合计", "结果": voucher_balance.get("credit_total", 0)},
        {"项目": "借贷平衡检查", "结果": "通过" if voucher_balance.get("balance_check") else "不通过"},
        {"项目": "复核报告", "结果": review_result.get("review_report", "")},
    ]


def write_sheet(writer: pd.ExcelWriter, sheet_name: str, rows: list[dict[str, Any]]) -> None:
    """
    写入一个工作表。
    """
    if rows:
        df = pd.DataFrame(rows)
    else:
        df = pd.DataFrame([{"提示": "无记录"}])

    df.to_excel(writer, sheet_name=sheet_name, index=False)


def append_review_exception_summary(ws: Worksheet, review_exceptions: list[dict[str, Any]]) -> None:
    """
    在“复核结论”工作表中追加完整异常清单摘要。
    一行一个异常点，避免复核报告只展示“其余异常详见异常清单”。
    """
    title_fill = PatternFill("solid", fgColor="BDD7EE")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin_side = Side(style="thin", color="D9D9D9")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    row = ws.max_row + 2

    # 分节标题
    ws.cell(row=row, column=1, value="四、异常清单摘要（完整明细）")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    ws.cell(row=row, column=1).fill = title_fill
    ws.cell(row=row, column=1).alignment = Alignment(vertical="center", wrap_text=True)
    ws.cell(row=row, column=1).border = border

    # 合并标题行，跨 4 列展示
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1

    if not review_exceptions:
        ws.cell(row=row, column=1, value="系统复核无异常，仍建议由财务人员人工最终确认。")
        ws.cell(row=row, column=1).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=row, column=1).border = border
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        return

    # 表头
    headers = ["序号", "异常类型", "异常说明", "修正建议"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    row += 1

    # 明细行
    for idx, item in enumerate(review_exceptions, start=1):
        exception_type = item.get("type", "")
        message = item.get("message", "")
        suggestion = item.get("suggestion", "")

        values = [idx, exception_type, message, suggestion]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border

        row += 1

    # 说明
    row += 1
    ws.cell(row=row, column=1, value="说明")
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=1).fill = header_fill
    ws.cell(row=row, column=1).alignment = Alignment(vertical="center", wrap_text=True)
    ws.cell(row=row, column=1).border = border

    ws.cell(
        row=row,
        column=2,
        value="以上异常清单为系统逐项比对后的完整摘要。请结合“复核异常清单”工作表逐项核对并修正；修正后可重新上传复核。"
    )
    ws.cell(row=row, column=2).alignment = Alignment(vertical="top", wrap_text=True)
    ws.cell(row=row, column=2).border = border
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)

    # 固定关键列宽，便于查看长文本
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 95
    ws.column_dimensions["D"].width = 65


def style_review_summary_sheet(ws: Worksheet) -> None:
    """
    对“复核结论”工作表做额外样式优化。
    """
    # 复核报告所在的“结果”列较长，给 B 列更大宽度
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 95
    ws.column_dimensions["C"].width = 95
    ws.column_dimensions["D"].width = 65

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # 适当提高行高，避免长文本完全挤在一起
    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 36


def generate_review_report_excel(
    review_result: dict[str, Any],
    output_dir: Path,
    run_id: str
) -> dict[str, Any]:
    """
    生成复核报告 Excel。
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"党费复核报告_{run_id}.xlsx"

    summary_rows = build_summary_rows(review_result)
    review_exception_rows = normalize_records(review_result.get("review_exceptions", []) or [])
    original_exception_rows = normalize_records(review_result.get("original_business_exceptions", []) or [])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        write_sheet(writer, "复核结论", summary_rows)
        write_sheet(writer, "复核异常清单", review_exception_rows)
        write_sheet(writer, "原始业务异常", original_exception_rows)

    wb = load_workbook(output_path)

    # 先追加完整异常摘要到“复核结论”sheet
    if "复核结论" in wb.sheetnames:
        ws_summary = wb["复核结论"]
        append_review_exception_summary(ws_summary, review_exception_rows)

    # 统一美化所有 sheet
    for ws in wb.worksheets:
        style_worksheet(ws)

    # 对“复核结论”sheet 做二次优化，保留异常摘要的大列宽
    if "复核结论" in wb.sheetnames:
        style_review_summary_sheet(wb["复核结论"])

    wb.save(output_path)

    return {
        "review_report_file_name": output_path.name,
        "review_report_file_path": str(output_path),
        "review_report_sheet_count": 3,
    }