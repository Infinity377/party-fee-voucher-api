from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import FileResponse
from pathlib import Path
from typing import Any
from urllib.parse import quote
import os
import shutil
import uuid
import pandas as pd

from utils.cleaners import (
    clean_bank_flow,
    clean_oa_flow,
    clean_subject_table,
    clean_member_status_table,
    clean_rule_mapping_table,
    dataframe_to_records,
)

from utils.matcher import match_all_business
from utils.voucher_generator import generate_voucher_excel
from utils.ledger_generator import generate_ledger_excel
from utils.review_checker import perform_review
from utils.review_report_generator import generate_review_report_excel
from utils.ai_suggestion_applier import apply_ai_suggestions_to_drafts

app = FastAPI(title="党费凭证Dify Python服务")

BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"
TEMPLATE_DIR = BASE_DIR / "templates"
TEMP_DIR = BASE_DIR / "temp_uploads"
OUTPUT_DIR = BASE_DIR / "output"

TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

SUBJECT_FILE = CONFIG_DIR / "会计科目表_党.xlsx"
MEMBER_STATUS_FILE = CONFIG_DIR / "党员离退休情况表.xlsx"
RULE_MAPPING_FILE = CONFIG_DIR / "党费业务映射规则表.xlsx"

VOUCHER_TEMPLATE_FILE = TEMPLATE_DIR / "凭证模板.xlsx"
LEDGER_TEMPLATE_FILE = TEMPLATE_DIR / "台账模板.xlsx"


def safe_filename(upload_file: UploadFile, default_name: str) -> str:
    """
    处理 UploadFile.filename 可能为空的问题。
    """
    if upload_file.filename and upload_file.filename.strip():
        return Path(upload_file.filename).name
    return default_name


def save_upload_file(upload_file: UploadFile, target_dir: Path, default_name: str) -> Path:
    """
    保存上传文件，并返回保存路径。
    """
    filename = safe_filename(upload_file, default_name)
    save_path = target_dir / filename

    with save_path.open("wb") as f:
        shutil.copyfileobj(upload_file.file, f)

    return save_path


def normalize_value(value: Any) -> Any:
    """
    将 pandas / Excel 中不适合 JSON 输出的值转成普通 Python 类型。
    """
    if pd.isna(value):
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def dataframe_preview(df: pd.DataFrame, max_rows: int = 5) -> dict[str, Any]:
    """
    返回 DataFrame 的字段、行数、前几行样例。
    """
    preview_df = df.head(max_rows).copy()

    rows = []
    for _, row in preview_df.iterrows():
        rows.append({
            str(col): normalize_value(row[col])
            for col in preview_df.columns
        })

    return {
        "row_count": int(len(df)),
        "column_count": int(len(df.columns)),
        "columns": [str(col) for col in df.columns],
        "preview_rows": rows
    }


def read_voucher(file_path: Path) -> pd.DataFrame:
    """
    读取凭证文件。
    """
    df = pd.read_excel(file_path, sheet_name=0, header=1)
    df = df.dropna(how="all")
    df = df.drop(
        columns=[col for col in df.columns if "billhead_" in str(col)],
        errors="ignore"
    )
    return df


def read_ledger(file_path: Path) -> pd.DataFrame:
    """
    读取台账文件。
    """
    df = pd.read_excel(file_path, sheet_name=0, header=0)
    df = df.dropna(how="all")
    return df


def read_config_files() -> dict[str, Any]:
    """
    读取系统配置文件。
    """
    result: dict[str, Any] = {}

    config_targets = {
        "subject_table": SUBJECT_FILE,
        "member_status_table": MEMBER_STATUS_FILE,
        "rule_mapping_table": RULE_MAPPING_FILE
    }

    for key, path in config_targets.items():
        if not path.exists():
            result[key] = {
                "exists": False,
                "path": str(path),
                "error": "文件不存在"
            }
            continue

        try:
            df = pd.read_excel(path, sheet_name=0, header=0)
            df = df.dropna(how="all")
            df = df.dropna(axis=1, how="all")

            result[key] = {
                "exists": True,
                "path": str(path),
                "preview": dataframe_preview(df)
            }
        except Exception as exc:
            result[key] = {
                "exists": True,
                "path": str(path),
                "error": str(exc)
            }

    return result


def check_template_files() -> dict[str, Any]:
    """
    检查模板文件是否存在。
    """
    return {
        "voucher_template": {
            "exists": VOUCHER_TEMPLATE_FILE.exists(),
            "path": str(VOUCHER_TEMPLATE_FILE)
        },
        "ledger_template": {
            "exists": LEDGER_TEMPLATE_FILE.exists(),
            "path": str(LEDGER_TEMPLATE_FILE)
        }
    }


def build_download_url(request: Request, file_name: str) -> str:
    """
    生成给用户点击的下载链接。

    本地默认使用 127.0.0.1；
    云端部署时通过环境变量 PUBLIC_BASE_URL 指定公网域名。
    """
    public_base_url = os.getenv("PUBLIC_BASE_URL", "http://127.0.0.1:8000").rstrip("/")
    encoded_file_name = quote(file_name)
    return f"{public_base_url}/download/{encoded_file_name}"


@app.get("/")
def health_check():
    return {
        "status": "ok",
        "message": "党费凭证Dify Python服务已启动"
    }


@app.get("/inspect-config")
def inspect_config():
    """
    单独检查系统配置和模板是否准备好。
    """
    return {
        "status": "success",
        "config_files": read_config_files(),
        "template_files": check_template_files()
    }


@app.get("/download/{file_name}")
def download_output_file(file_name: str):
    """
    下载 output 文件夹中的生成结果文件。
    """
    safe_name = Path(file_name).name
    file_path = OUTPUT_DIR / safe_name

    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail=f"文件不存在：{safe_name}")

    return FileResponse(
        path=file_path,
        filename=safe_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/generate")
async def generate_voucher_and_ledger(
    request: Request,
    bank_flow_file: UploadFile = File(...),
    oa_flow_file: UploadFile = File(...),
    maker: str = Form(default=""),
    book_code: str = Form(default=""),
    voucher_type: str = Form(default=""),
    business_note: str = Form(default="")
):
    """
    入口一：生成凭证草稿和台账草稿。
    """
    run_id = str(uuid.uuid4())[:8]
    run_dir = TEMP_DIR / f"generate_{run_id}"
    run_dir.mkdir(parents=True, exist_ok=True)

    bank_path = save_upload_file(bank_flow_file, run_dir, "bank_flow.xlsx")
    oa_path = save_upload_file(oa_flow_file, run_dir, "oa_flow.xlsx")

    try:
        bank_df = clean_bank_flow(bank_path)
        oa_df = clean_oa_flow(oa_path)
        subject_df = clean_subject_table(SUBJECT_FILE)
        member_df = clean_member_status_table(MEMBER_STATUS_FILE)
        rule_df = clean_rule_mapping_table(RULE_MAPPING_FILE)

        direction_counts = bank_df["direction"].value_counts(dropna=False).to_dict()

        matched_df, business_exceptions = match_all_business(
            bank_df=bank_df,
            oa_df=oa_df,
            subject_df=subject_df,
            member_df=member_df,
            rule_df=rule_df
        )

        matched_count = int((matched_df["match_status"] == "matched").sum())
        exception_count = int(len(business_exceptions))

        voucher_result = generate_voucher_excel(
            matched_df=matched_df,
            template_path=VOUCHER_TEMPLATE_FILE,
            output_dir=OUTPUT_DIR,
            run_id=run_id,
            maker=maker,
            book_code=book_code,
            voucher_type=voucher_type
        )

        ledger_result = generate_ledger_excel(
            matched_df=matched_df,
            template_path=LEDGER_TEMPLATE_FILE,
            output_dir=OUTPUT_DIR,
            run_id=run_id
        )

        voucher_download_url = build_download_url(
            request=request,
            file_name=voucher_result["voucher_file_name"]
        )
        ledger_download_url = build_download_url(
            request=request,
            file_name=ledger_result["ledger_file_name"]
        )

        return {
            "status": "success",
            "mode": "generate",
            "message": "已完成流水、OA、规则表清洗、业务匹配，并生成凭证草稿 Excel 和台账草稿 Excel。",
            "run_id": run_id,
            "received_files": {
                "bank_flow_file": bank_path.name,
                "oa_flow_file": oa_path.name
            },
            "params": {
                "maker": maker,
                "book_code": book_code,
                "voucher_type": voucher_type,
                "business_note": business_note
            },
            "summary": {
                "bank_flow_rows": int(len(bank_df)),
                "oa_flow_rows": int(len(oa_df)),
                "subject_rows": int(len(subject_df)),
                "member_rows": int(len(member_df)),
                "rule_rows": int(len(rule_df)),
                "bank_direction_counts": direction_counts
            },
            "business_match_summary": {
                "matched_count": matched_count,
                "exception_count": exception_count
            },
            "voucher_result": voucher_result,
            "ledger_result": ledger_result,
            "download_links": {
                "voucher_download_url": voucher_download_url,
                "ledger_download_url": ledger_download_url
            },
            "matched_business_columns": list(matched_df.columns),
            "matched_business_preview": dataframe_to_records(matched_df, max_rows=30),
            "cleaned_bank_flow_columns": list(bank_df.columns),
            "cleaned_bank_flow_preview": dataframe_to_records(bank_df, max_rows=20),
            "cleaned_oa_flow_columns": list(oa_df.columns),
            "cleaned_oa_flow_preview": dataframe_to_records(oa_df, max_rows=20),
            "cleaned_subject_columns": list(subject_df.columns),
            "cleaned_subject_preview": dataframe_to_records(subject_df, max_rows=20),
            "cleaned_member_columns": list(member_df.columns),
            "cleaned_member_preview": dataframe_to_records(member_df, max_rows=20),
            "cleaned_rule_columns": list(rule_df.columns),
            "cleaned_rule_preview": dataframe_to_records(rule_df, max_rows=20),
            "config_files": read_config_files(),
            "template_files": check_template_files(),
            "has_exception": exception_count > 0,
            "exceptions": business_exceptions
        }

    except Exception as exc:
        return {
            "status": "error",
            "mode": "generate",
            "message": "读取、清洗、业务匹配、凭证生成或台账生成失败",
            "run_id": run_id,
            "error": str(exc),
            "has_exception": True,
            "exceptions": [
                {
                    "type": "GENERATE_ERROR",
                    "message": str(exc)
                }
            ]
        }

@app.post("/review")
async def review_voucher_and_ledger(
    request: Request,
    bank_flow_file: UploadFile = File(...),
    oa_flow_file: UploadFile = File(...),
    voucher_draft_file: UploadFile = File(...),
    ledger_draft_file: UploadFile = File(...),
    maker: str = Form(default=""),
    book_code: str = Form(default=""),
    voucher_type: str = Form(default=""),
    review_note: str = Form(default="")
):

    """
    入口二：复核凭证草稿和台账草稿。
    当前版本：清洗流水、OA、配置表，并读取凭证和台账结构。
    """
    run_id = str(uuid.uuid4())[:8]
    run_dir = TEMP_DIR / f"review_{run_id}"
    run_dir.mkdir(parents=True, exist_ok=True)

    bank_path = save_upload_file(bank_flow_file, run_dir, "bank_flow.xlsx")
    oa_path = save_upload_file(oa_flow_file, run_dir, "oa_flow.xlsx")
    voucher_path = save_upload_file(voucher_draft_file, run_dir, "voucher_draft.xlsx")
    ledger_path = save_upload_file(ledger_draft_file, run_dir, "ledger_draft.xlsx")

    try:
        bank_df = clean_bank_flow(bank_path)
        oa_df = clean_oa_flow(oa_path)
        subject_df = clean_subject_table(SUBJECT_FILE)
        member_df = clean_member_status_table(MEMBER_STATUS_FILE)
        rule_df = clean_rule_mapping_table(RULE_MAPPING_FILE)

        voucher_df = read_voucher(voucher_path)
        ledger_df = read_ledger(ledger_path)

        review_result = perform_review(
            bank_df=bank_df,
            oa_df=oa_df,
            subject_df=subject_df,
            member_df=member_df,
            rule_df=rule_df,
            voucher_path=voucher_path,
            ledger_path=ledger_path,
            maker=maker,
            book_code=book_code,
            voucher_type=voucher_type
        )

        review_report_result = generate_review_report_excel(
            review_result=review_result,
            output_dir=OUTPUT_DIR,
            run_id=run_id
        )

        review_report_download_url = build_download_url(
            request=request,
            file_name=review_report_result["review_report_file_name"]
        )

        return {
            "status": "success",
            "mode": "review",
            "message": "已完成凭证草稿和台账草稿的系统复核。",
            "run_id": run_id,
            "received_files": {
                "bank_flow_file": bank_path.name,
                "oa_flow_file": oa_path.name,
                "voucher_draft_file": voucher_path.name,
                "ledger_draft_file": ledger_path.name
            },
            "params": {
                "maker": maker,
                "book_code": book_code,
                "voucher_type": voucher_type,
                "review_note": review_note
            },
            "summary": {
                "bank_flow_rows": int(len(bank_df)),
                "oa_flow_rows": int(len(oa_df)),
                "subject_rows": int(len(subject_df)),
                "member_rows": int(len(member_df)),
                "rule_rows": int(len(rule_df)),
                "voucher_rows": int(len(voucher_df)),
                "ledger_rows": int(len(ledger_df)),
                "bank_direction_counts": bank_df["direction"].value_counts(dropna=False).to_dict()
            },
            "review_result": review_result,
            "review_report_result": review_report_result,
            "download_links": {
                "review_report_download_url": review_report_download_url
            },
            "cleaned_bank_flow_preview": dataframe_to_records(bank_df, max_rows=20),
            "cleaned_oa_flow_preview": dataframe_to_records(oa_df, max_rows=20),
            "voucher_columns": [str(col) for col in voucher_df.columns],
            "voucher_preview": dataframe_preview(voucher_df),
            "ledger_columns": [str(col) for col in ledger_df.columns],
            "ledger_preview": dataframe_preview(ledger_df),
            "has_exception": not review_result["review_passed"],
            "exceptions": review_result["review_exceptions"],
            "review_report": review_result["review_report"]
        }

    except Exception as exc:
        return {
            "status": "error",
            "mode": "review",
            "message": "读取或清洗Excel失败",
            "run_id": run_id,
            "error": str(exc),
            "has_exception": True,
            "exceptions": [
                {
                    "type": "EXCEL_CLEAN_ERROR",
                    "message": str(exc)
                }
            ]
        }
    
    
@app.post("/apply_ai_suggestions")
async def apply_ai_suggestions(
    request: Request,
    voucher_file_name: str = Form(...),
    ledger_file_name: str = Form(...),
    ai_suggestions_json: str = Form(...),
    maker: str = Form("何家俊"),
    book_code: str = Form("501-0007"),
    voucher_type: str = Form("01"),
    note: str = Form("根据AI建议补录异常业务"),
) -> dict[str, Any]:
    """
    根据 Dify LLM1 输出的异常业务 AI 补录建议，
    在已生成的凭证草稿和台账草稿基础上追加 AI 建议分录，
    生成“AI补录版”凭证和台账。

    注意：
    1. 不覆盖原始凭证草稿和台账草稿；
    2. 候选科目必须存在于会计科目表；
    3. AI 补录分录仅为待人工复核版本。
    """
    try:
        run_id = uuid.uuid4().hex[:8]

        safe_voucher_name = Path(voucher_file_name).name
        safe_ledger_name = Path(ledger_file_name).name

        voucher_file_path = OUTPUT_DIR / safe_voucher_name
        ledger_file_path = OUTPUT_DIR / safe_ledger_name

        result = apply_ai_suggestions_to_drafts(
            voucher_file_path=voucher_file_path,
            ledger_file_path=ledger_file_path,
            output_dir=OUTPUT_DIR,
            run_id=run_id,
            ai_suggestions_json=ai_suggestions_json,
            subject_file=SUBJECT_FILE,
            maker=maker,
            book_code=book_code,
            voucher_type=voucher_type,
        )

        modified_voucher_file_name = result.get("modified_voucher_file_name", "")
        modified_ledger_file_name = result.get("modified_ledger_file_name", "")

        modified_voucher_download_url = build_download_url(request, modified_voucher_file_name)
        modified_ledger_download_url = build_download_url(request, modified_ledger_file_name)

        applied_count = result.get("applied_count", 0)
        skipped_count = result.get("skipped_count", 0)

        if applied_count > 0:
            message = f"已根据AI建议生成补录版凭证和台账，本次成功补录 {applied_count} 笔，跳过 {skipped_count} 笔。"
        else:
            message = f"未写入AI补录分录，已生成补录版文件副本；跳过 {skipped_count} 笔建议。"

        return {
            "status": "success",
            "mode": "apply_ai_suggestions",
            "message": message,
            "run_id": run_id,
            "received_files": {
                "voucher_file_name": safe_voucher_name,
                "ledger_file_name": safe_ledger_name,
            },
            "params": {
                "maker": maker,
                "book_code": book_code,
                "voucher_type": voucher_type,
                "note": note,
            },
            "apply_result": result,
            "download_links": {
                "modified_voucher_download_url": modified_voucher_download_url,
                "modified_ledger_download_url": modified_ledger_download_url,
            },
        }

    except Exception as e:
        return {
            "status": "error",
            "mode": "apply_ai_suggestions",
            "message": "根据AI建议补录凭证和台账失败",
            "error": str(e),
            "has_exception": True,
        }