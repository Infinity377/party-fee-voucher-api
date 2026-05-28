from __future__ import annotations

import json
import re
import shutil
from calendar import monthrange
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook


def _strip_code_fence(text: str) -> str:
    """
    兼容 LLM 输出 ```json ... ``` 的情况。
    """
    text = (text or "").strip()

    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text.strip(), flags=re.IGNORECASE).strip()
        text = re.sub(r"```$", "", text.strip()).strip()

    return text


def _to_float(value: Any) -> float:
    if value is None:
        return 0.0

    text = str(value).replace(",", "").strip()
    if text == "":
        return 0.0

    return float(text)


def _clean_code(value: Any) -> str:
    """
    统一清洗科目编码，避免 Excel 把 400101 读成 400101.0。
    """
    if value is None:
        return ""

    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]

    return text


def _normalize_date(value: Any) -> str:
    """
    将日期统一为 YYYY-MM-DD。
    """
    if value is None:
        raise ValueError("交易日期为空")

    text = str(value).strip()
    if not text:
        raise ValueError("交易日期为空")

    # 兼容 2026/03/17
    text = text.replace("/", "-")

    # 兼容 2026-03-17 09:00:00
    if " " in text:
        text = text.split(" ")[0]

    parts = text.split("-")
    if len(parts) != 3:
        raise ValueError(f"无法识别交易日期：{value}")

    year = int(parts[0])
    month = int(parts[1])
    day = int(parts[2])

    return f"{year:04d}-{month:02d}-{day:02d}"


def _month_end(date_text: str) -> str:
    year, month, _ = [int(x) for x in date_text.split("-")]
    last_day = monthrange(year, month)[1]
    return f"{year:04d}-{month:02d}-{last_day:02d}"


def _parse_ai_suggestions(ai_suggestions_json: str) -> list[dict[str, Any]]:
    """
    解析 LLM1 输出的 JSON。
    支持两种格式：
    1. {"suggestions": [...]}
    2. [...]
    """
    raw = _strip_code_fence(ai_suggestions_json)

    if not raw:
        return []

    data = json.loads(raw)

    if isinstance(data, list):
        suggestions = data
    elif isinstance(data, dict):
        suggestions = data.get("suggestions", []) or []
    else:
        suggestions = []

    if not isinstance(suggestions, list):
        return []

    normalized: list[dict[str, Any]] = []
    for item in suggestions:
        if isinstance(item, dict):
            normalized.append(item)

    return normalized


def _load_valid_subject_codes(subject_file: Path) -> set[str]:
    """
    从会计科目表中读取合法科目编码。
    只允许 AI 建议使用会计科目表中存在的科目，防止 LLM 编造科目。
    """
    if not subject_file.exists():
        return set()

    df = pd.read_excel(subject_file)

    code_col = None
    for col in df.columns:
        col_text = str(col)
        if "科目编码" in col_text or col_text.strip() in {"编码", "科目"}:
            code_col = col
            break

    if code_col is None:
        # 兜底：找第一个包含“编码”的列
        for col in df.columns:
            if "编码" in str(col):
                code_col = col
                break

    if code_col is None:
        return set()

    codes: set[str] = set()
    for value in df[code_col].dropna().tolist():
        code = _clean_code(value)
        if code:
            codes.add(code)

    return codes


def _get_header_map(ws) -> dict[str, int]:
    """
    读取第 1 行表头，返回列名 -> 列号。
    """
    header_map: dict[str, int] = {}

    for cell in ws[1]:
        if cell.value is None:
            continue
        header_map[str(cell.value).strip()] = int(cell.column)

    return header_map


def _set_if_exists(ws, row: int, header_map: dict[str, int], col_name: str, value: Any) -> None:
    col = header_map.get(col_name)
    if col is not None:
        ws.cell(row=row, column=col, value=value)


def _get_next_custom_no_for_month(ws, header_map: dict[str, int], voucher_date: str) -> int:
    """
    表头自定义项3：按月重新编号。
    读取现有凭证中同一制单月份的最大编号，然后 +1。
    """
    date_col = header_map.get("* 制单日期")
    custom_col = header_map.get("表头自定义项3")

    if date_col is None or custom_col is None:
        return 1

    month_key = voucher_date[:7]
    max_no = 0

    for row in range(2, ws.max_row + 1):
        date_value = ws.cell(row=row, column=date_col).value
        custom_value = ws.cell(row=row, column=custom_col).value

        if date_value is None or custom_value is None:
            continue

        date_text = str(date_value).strip()
        if date_text.startswith(month_key):
            try:
                max_no = max(max_no, int(float(str(custom_value))))
            except Exception:
                continue

    return max_no + 1


def _append_voucher_rows(
    ws,
    suggestion: dict[str, Any],
    header_map: dict[str, int],
    maker: str,
    book_code: str,
    voucher_type: str,
) -> None:
    """
    根据一条 AI 建议追加凭证借贷两行。
    收入：借 1002，贷 候选科目
    支出：借 候选科目，贷 1002
    """
    direction = str(suggestion.get("direction", "")).strip()
    transaction_date = _normalize_date(suggestion.get("transaction_date"))
    voucher_date = _month_end(transaction_date)

    amount = _to_float(suggestion.get("amount"))
    subject_code = _clean_code(suggestion.get("candidate_subject_code"))
    summary = str(suggestion.get("suggested_summary", "")).strip()

    if not summary:
        counterparty = str(suggestion.get("counterparty", "")).strip()
        if direction == "收入":
            summary = f"收到{counterparty}划来党费"
        else:
            summary = f"支付{counterparty}相关党费业务款项"

    custom_no = _get_next_custom_no_for_month(ws, header_map, voucher_date)

    debit_row = ws.max_row + 1
    credit_row = ws.max_row + 2

    common_values = {
        "* 核算账簿": book_code,
        "* 凭证类别": voucher_type,
        "* 凭证号": "",
        "附单据数": "",
        "* 制单人": maker,
        "* 制单日期": voucher_date,
        "审核人": "",
        "审核日期": "",
        "* 摘要": summary,
        "表头自定义项2": "AI建议补录-待复核",
        "表头自定义项3": custom_no,
        "* 币种": "CNY",
        "结算号": "",
        "结算日期": transaction_date,
        "结算方式": "",
        "核销号": "",
        "核销业务日期": transaction_date,
    }

    for row in [debit_row, credit_row]:
        for col_name, value in common_values.items():
            _set_if_exists(ws, row, header_map, col_name, value)

    if direction == "收入":
        # 借：银行存款 1002
        _set_if_exists(ws, debit_row, header_map, "* 科目编码", "1002")
        _set_if_exists(ws, debit_row, header_map, "* 原币借方金额", amount)
        _set_if_exists(ws, debit_row, header_map, "* 本币借方金额", amount)
        _set_if_exists(ws, debit_row, header_map, "* 原币贷方金额", "")
        _set_if_exists(ws, debit_row, header_map, "* 本币贷方金额", "")

        # 贷：候选收入科目
        _set_if_exists(ws, credit_row, header_map, "* 科目编码", subject_code)
        _set_if_exists(ws, credit_row, header_map, "* 原币借方金额", "")
        _set_if_exists(ws, credit_row, header_map, "* 本币借方金额", "")
        _set_if_exists(ws, credit_row, header_map, "* 原币贷方金额", amount)
        _set_if_exists(ws, credit_row, header_map, "* 本币贷方金额", amount)

    elif direction == "支出":
        # 借：候选支出科目
        _set_if_exists(ws, debit_row, header_map, "* 科目编码", subject_code)
        _set_if_exists(ws, debit_row, header_map, "* 原币借方金额", amount)
        _set_if_exists(ws, debit_row, header_map, "* 本币借方金额", amount)
        _set_if_exists(ws, debit_row, header_map, "* 原币贷方金额", "")
        _set_if_exists(ws, debit_row, header_map, "* 本币贷方金额", "")

        # 贷：银行存款 1002
        _set_if_exists(ws, credit_row, header_map, "* 科目编码", "1002")
        _set_if_exists(ws, credit_row, header_map, "* 原币借方金额", "")
        _set_if_exists(ws, credit_row, header_map, "* 本币借方金额", "")
        _set_if_exists(ws, credit_row, header_map, "* 原币贷方金额", amount)
        _set_if_exists(ws, credit_row, header_map, "* 本币贷方金额", amount)

    else:
        raise ValueError(f"无法识别AI建议方向：{direction}")


def _append_ledger_row(ws, suggestion: dict[str, Any], header_map: dict[str, int]) -> None:
    """
    根据一条 AI 建议追加台账一行。
    """
    direction = str(suggestion.get("direction", "")).strip()
    transaction_date = _normalize_date(suggestion.get("transaction_date"))
    voucher_date = _month_end(transaction_date)

    year, month, day = [int(x) for x in voucher_date.split("-")]

    amount = _to_float(suggestion.get("amount"))
    summary = str(suggestion.get("suggested_summary", "")).strip()
    ledger_tag = str(suggestion.get("ledger_tag", "")).strip()

    if not ledger_tag:
        ledger_tag = "支出" if direction == "支出" else "公司党员统一上缴"

    row = ws.max_row + 1

    _set_if_exists(ws, row, header_map, "年", year)
    _set_if_exists(ws, row, header_map, "月", month)
    _set_if_exists(ws, row, header_map, "日", day)
    _set_if_exists(ws, row, header_map, "编号", "")
    _set_if_exists(ws, row, header_map, "摘要", summary)

    if direction == "收入":
        _set_if_exists(ws, row, header_map, "支出", "")
        _set_if_exists(ws, row, header_map, "收入", amount)
    elif direction == "支出":
        _set_if_exists(ws, row, header_map, "支出", amount)
        _set_if_exists(ws, row, header_map, "收入", "")
    else:
        raise ValueError(f"无法识别AI建议方向：{direction}")

    _set_if_exists(ws, row, header_map, "余额（元）", "")
    _set_if_exists(ws, row, header_map, "标签", ledger_tag)


def _validate_suggestion(
    suggestion: dict[str, Any],
    valid_subject_codes: set[str],
) -> tuple[bool, str]:
    """
    校验 AI 建议是否可写入草稿。
    """
    subject_code = _clean_code(suggestion.get("candidate_subject_code"))
    direction = str(suggestion.get("direction", "")).strip()
    amount = _to_float(suggestion.get("amount"))

    if direction not in {"收入", "支出"}:
        return False, f"方向不是收入/支出：{direction}"

    if amount <= 0:
        return False, f"金额必须大于0：{amount}"

    if not subject_code:
        return False, "候选科目编码为空"

    if subject_code == "1002":
        return False, "候选科目不能是银行存款1002"

    if valid_subject_codes and subject_code not in valid_subject_codes:
        return False, f"候选科目编码不在会计科目表中：{subject_code}"

    try:
        _normalize_date(suggestion.get("transaction_date"))
    except Exception as exc:
        return False, str(exc)

    return True, ""


def apply_ai_suggestions_to_drafts(
    voucher_file_path: Path,
    ledger_file_path: Path,
    output_dir: Path,
    run_id: str,
    ai_suggestions_json: str,
    subject_file: Path,
    maker: str,
    book_code: str,
    voucher_type: str,
) -> dict[str, Any]:
    """
    根据 LLM1 输出的 AI 补录建议，生成“AI补录版”凭证草稿和台账草稿。

    设计原则：
    1. 不覆盖原始草稿；
    2. 只生成新增的 AI 补录版文件；
    3. 候选科目必须在会计科目表中；
    4. 追加行统一标注“AI建议补录-待复核”；
    5. 所有 AI 补录结果仍需财务人员人工复核。
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    if not voucher_file_path.exists():
        raise FileNotFoundError(f"凭证草稿文件不存在：{voucher_file_path}")

    if not ledger_file_path.exists():
        raise FileNotFoundError(f"台账草稿文件不存在：{ledger_file_path}")

    suggestions = _parse_ai_suggestions(ai_suggestions_json)
    valid_subject_codes = _load_valid_subject_codes(subject_file)

    modified_voucher_path = output_dir / f"{voucher_file_path.stem}_AI补录版_{run_id}.xlsx"
    modified_ledger_path = output_dir / f"{ledger_file_path.stem}_AI补录版_{run_id}.xlsx"

    shutil.copy2(voucher_file_path, modified_voucher_path)
    shutil.copy2(ledger_file_path, modified_ledger_path)

    applied: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []

    voucher_wb = load_workbook(modified_voucher_path)
    ledger_wb = load_workbook(modified_ledger_path)

    voucher_ws = voucher_wb.active
    ledger_ws = ledger_wb.active

    voucher_header_map = _get_header_map(voucher_ws)
    ledger_header_map = _get_header_map(ledger_ws)

    for suggestion in suggestions:
        should_apply = suggestion.get("apply_to_draft", True)
        if should_apply is False:
            skipped.append({
                "suggestion": suggestion,
                "reason": "LLM建议不写入草稿"
            })
            continue

        ok, reason = _validate_suggestion(suggestion, valid_subject_codes)
        if not ok:
            skipped.append({
                "suggestion": suggestion,
                "reason": reason
            })
            continue

        try:
            _append_voucher_rows(
                ws=voucher_ws,
                suggestion=suggestion,
                header_map=voucher_header_map,
                maker=maker,
                book_code=book_code,
                voucher_type=voucher_type,
            )
            _append_ledger_row(
                ws=ledger_ws,
                suggestion=suggestion,
                header_map=ledger_header_map,
            )

            applied.append({
                "flow_index": suggestion.get("flow_index", ""),
                "direction": suggestion.get("direction", ""),
                "transaction_date": suggestion.get("transaction_date", ""),
                "counterparty": suggestion.get("counterparty", ""),
                "amount": suggestion.get("amount", ""),
                "candidate_subject_code": _clean_code(suggestion.get("candidate_subject_code")),
                "candidate_subject_name": suggestion.get("candidate_subject_name", ""),
                "suggested_summary": suggestion.get("suggested_summary", ""),
                "ledger_tag": suggestion.get("ledger_tag", ""),
                "confidence": suggestion.get("confidence", ""),
                "review_prompt": suggestion.get("review_prompt", ""),
            })

        except Exception as exc:
            skipped.append({
                "suggestion": suggestion,
                "reason": f"写入失败：{str(exc)}"
            })

    voucher_wb.save(modified_voucher_path)
    ledger_wb.save(modified_ledger_path)

    return {
        "applied_count": len(applied),
        "skipped_count": len(skipped),
        "applied_suggestions": applied,
        "skipped_suggestions": skipped,
        "modified_voucher_file_name": modified_voucher_path.name,
        "modified_voucher_file_path": str(modified_voucher_path),
        "modified_ledger_file_name": modified_ledger_path.name,
        "modified_ledger_file_path": str(modified_ledger_path),
    }