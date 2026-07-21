# -*- coding: utf-8 -*-
from __future__ import annotations

import asyncio
import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

from extractor import (
    ALLOWED_COLUMNS,
    apply_custom_fields,
    canonical_column_name,
    get_row_value,
    parse_pdf,
    write_excel,
)


APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
MAX_FILE_COUNT = int(os.getenv("MAX_FILE_COUNT", "500"))
MAX_UPLOAD_SIZE = int(os.getenv("MAX_UPLOAD_SIZE_MB", "25")) * 1024 * 1024
MAX_TOTAL_UPLOAD_SIZE = int(os.getenv("MAX_TOTAL_UPLOAD_SIZE_MB", "5120")) * 1024 * 1024
UPLOAD_CHUNK_SIZE = int(os.getenv("UPLOAD_CHUNK_SIZE_MB", "1")) * 1024 * 1024
PARSE_CONCURRENCY = max(1, int(os.getenv("PARSE_CONCURRENCY", str(min(4, os.cpu_count() or 1)))))
PREVIEW_ROW_LIMIT = int(os.getenv("PREVIEW_ROW_LIMIT", "500"))
MAX_CUSTOM_FIELDS = 60
MAX_EXPORT_COLUMNS = 80

app = FastAPI(title="PO Extractor Pro - 采购订单PDF智能提取工作台")


def _json_loads_list(raw: str | None) -> List[Any]:
    if not raw:
        return []
    try:
        value = json.loads(raw)
    except Exception:
        return []
    return value if isinstance(value, list) else []


def _safe_text(value: Any, max_len: int = 80) -> str:
    text = str(value or "").strip()
    text = text.replace("\x00", "")
    return text[:max_len].strip()


def parse_columns(columns_raw: str | None) -> List[str]:
    columns = _json_loads_list(columns_raw)
    result: List[str] = []
    for col in columns:
        name = canonical_column_name(str(col).strip())
        if name in ALLOWED_COLUMNS and name not in result:
            result.append(name)
    return result


def parse_custom_fields(custom_fields_raw: str | None) -> List[Dict[str, Any]]:
    fields = _json_loads_list(custom_fields_raw)
    result: List[Dict[str, Any]] = []
    allowed_modes = {"after_label", "next_line", "between_labels", "table_column", "regex"}
    allowed_scopes = {"item", "order"}

    for item in fields[:MAX_CUSTOM_FIELDS]:
        if not isinstance(item, dict):
            continue
        name = _safe_text(item.get("name"), 80)
        title = _safe_text(item.get("title") or item.get("display_name") or name, 80)
        label = _safe_text(item.get("label"), 120)
        mode = _safe_text(item.get("mode") or "after_label", 30)
        scope = _safe_text(item.get("scope") or "item", 20)
        end_label = _safe_text(item.get("end_label"), 120)
        pattern = _safe_text(item.get("pattern"), 500)
        value_type = _safe_text(item.get("value_type") or item.get("extract_type") or "raw", 30)
        if value_type not in {"raw", "number", "quantity", "unit", "number_unit", "quantity_unit"}:
            value_type = "raw"

        if not name or name.startswith("_"):
            continue
        if mode not in allowed_modes:
            mode = "after_label"
        if scope not in allowed_scopes:
            scope = "item"
        # 表格列模式用于“项目、物料、交货期、数量、价格、金额”等已经被解析器识别出来的列。
        if mode == "table_column" and not label:
            label = name
        if mode != "regex" and not label:
            continue
        if mode == "regex" and not pattern:
            continue

        result.append(
            {
                "name": name,
                "title": title or name,
                "label": label,
                "mode": mode,
                "scope": scope,
                "end_label": end_label,
                "pattern": pattern,
                "value_type": value_type,
            }
        )
    return result


def _normalise_order(value: Any, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def parse_export_columns(
    export_columns_raw: str | None,
    columns_raw: str | None,
    custom_fields_raw: str | None,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """解析新版导出列配置。

    新版每一列都包含：
    - key：内部数据来源字段，例如 Material、Project ref、custom_abc
    - title：Excel 表头显示名称，例如 物料、项目编号
    - order：Excel 中的显示顺序
    - type：system/custom
    - rule：自定义字段查找规则

    同时兼容旧版 columns + custom_fields。
    """
    raw_export_columns = _json_loads_list(export_columns_raw)
    export_defs: List[Dict[str, Any]] = []
    custom_rules: List[Dict[str, Any]] = []
    seen_keys: set[str] = set()

    if raw_export_columns:
        for idx, item in enumerate(raw_export_columns[:MAX_EXPORT_COLUMNS], start=1):
            if not isinstance(item, dict):
                continue
            col_type = _safe_text(item.get("type") or "system", 20)
            raw_key = _safe_text(item.get("key") or item.get("name") or item.get("title"), 100)
            title = _safe_text(item.get("title") or item.get("name") or raw_key, 80)
            order = _normalise_order(item.get("order"), idx)

            if not raw_key or raw_key.startswith("_") or not title:
                continue

            if col_type == "custom":
                key = raw_key
                rule_obj = dict(item.get("rule") or {})
                # 兼容把规则字段放在顶层的写法。
                for rule_key in ("label", "mode", "scope", "end_label", "pattern", "value_type", "extract_type"):
                    if rule_key in item and rule_key not in rule_obj:
                        rule_obj[rule_key] = item[rule_key]
                rule_obj["name"] = key
                rule_obj["title"] = title
                parsed_rules = parse_custom_fields(json.dumps([rule_obj], ensure_ascii=False))
                if not parsed_rules:
                    continue
                custom_rules.append(parsed_rules[0])
            else:
                key = canonical_column_name(raw_key)
                if key not in ALLOWED_COLUMNS:
                    continue
                col_type = "system"

            if key in seen_keys:
                continue
            seen_keys.add(key)
            export_defs.append({"key": key, "title": title, "order": order, "type": col_type})

        export_defs.sort(key=lambda x: x.get("order", 0))
        # 重新写 order，避免前端传 1、5、99 时 Excel 中出现空列。
        for idx, col in enumerate(export_defs, start=1):
            col["order"] = idx
        return export_defs, custom_rules

    # 旧版兼容：columns + custom_fields。
    selected_columns = parse_columns(columns_raw)
    custom_rules = parse_custom_fields(custom_fields_raw)
    for idx, name in enumerate(selected_columns, start=1):
        if name not in seen_keys:
            seen_keys.add(name)
            export_defs.append({"key": name, "title": name, "order": idx, "type": "system"})
    for rule in custom_rules:
        key = str(rule.get("name") or "").strip()
        if key and not key.startswith("_") and key not in seen_keys:
            seen_keys.add(key)
            export_defs.append({"key": key, "title": key, "order": len(export_defs) + 1, "type": "custom"})
    return export_defs, custom_rules


async def save_uploads(files: List[UploadFile]):
    """流式保存上传文件，避免一次上传数百个 PDF 时占满内存。"""
    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_excel_"))
    pdf_paths: List[Path] = []
    errors: List[Dict[str, str]] = []
    total_size = 0

    accepted_files = files[:MAX_FILE_COUNT]
    for index, file in enumerate(accepted_files, start=1):
        original_name = Path(file.filename or "upload.pdf").name
        if not original_name.lower().endswith(".pdf"):
            errors.append({"file": original_name, "error": "已跳过：仅支持 PDF 文件。"})
            await file.close()
            continue

        safe_name = f"{index:04d}_{original_name}"
        target = temp_dir / safe_name
        file_size = 0
        header = b""

        try:
            with target.open("wb") as output:
                while True:
                    chunk = await file.read(UPLOAD_CHUNK_SIZE)
                    if not chunk:
                        break
                    if not header:
                        header = chunk[:4]
                    file_size += len(chunk)
                    total_size += len(chunk)

                    if file_size > MAX_UPLOAD_SIZE:
                        raise ValueError(f"文件超过 {MAX_UPLOAD_SIZE // 1024 // 1024}MB。")
                    if total_size > MAX_TOTAL_UPLOAD_SIZE:
                        raise OverflowError(
                            f"本次上传总大小超过 {MAX_TOTAL_UPLOAD_SIZE // 1024 // 1024}MB。"
                        )
                    output.write(chunk)

            if header != b"%PDF":
                target.unlink(missing_ok=True)
                errors.append({"file": original_name, "error": "已跳过：文件内容不是有效 PDF。"})
                continue

            pdf_paths.append(target)
        except ValueError as exc:
            target.unlink(missing_ok=True)
            errors.append({"file": original_name, "error": f"已跳过：{exc}"})
        except OverflowError:
            target.unlink(missing_ok=True)
            errors.append({
                "file": "全部文件",
                "error": f"本次上传总大小超过 {MAX_TOTAL_UPLOAD_SIZE // 1024 // 1024}MB，后续文件已停止接收。",
            })
            break
        finally:
            await file.close()

    if len(files) > MAX_FILE_COUNT:
        errors.append({
            "file": "全部文件",
            "error": f"一次最多处理 {MAX_FILE_COUNT} 个 PDF，其余 {len(files) - MAX_FILE_COUNT} 个已跳过。",
        })

    return temp_dir, pdf_paths, errors


async def parse_all_files(pdf_paths: Sequence[Path], custom_fields: Sequence[Dict[str, Any]]):
    """有限并发解析 PDF，提升数百文件批处理速度，同时防止 CPU/内存过载。"""
    semaphore = asyncio.Semaphore(PARSE_CONCURRENCY)

    async def parse_one(file_index: int, pdf_path: Path):
        display_name = (
            pdf_path.name[5:]
            if len(pdf_path.name) > 5 and pdf_path.name[:4].isdigit() and pdf_path.name[4] == "_"
            else pdf_path.name
        )
        async with semaphore:
            try:
                rows = await run_in_threadpool(parse_pdf, pdf_path)
                for row in rows:
                    row["来源文件"] = display_name
                    row["_source_file"] = display_name
                if custom_fields:
                    rows = await run_in_threadpool(apply_custom_fields, rows, custom_fields)
                if not rows:
                    return file_index, [], {
                        "file": display_name,
                        "error": "未提取到数据。请确认 PDF 是文字型，且属于当前采购订单格式。",
                    }
                return file_index, rows, None
            except Exception as exc:
                return file_index, [], {
                    "file": display_name,
                    "error": f"解析失败：{type(exc).__name__}。请确认 PDF 是文字型或降低解析并发后重试。",
                }

    results = await asyncio.gather(
        *(parse_one(index, path) for index, path in enumerate(pdf_paths))
    )
    results.sort(key=lambda item: item[0])

    all_rows: List[Dict[str, Any]] = []
    errors: List[Dict[str, str]] = []
    for _, rows, error in results:
        all_rows.extend(rows)
        if error:
            errors.append(error)
    return all_rows, errors


def filter_rows(rows: Sequence[Dict[str, Any]], export_defs: Sequence[Dict[str, Any]]):
    filtered = []
    for row in rows:
        out: Dict[str, Any] = {}
        used_titles: set[str] = set()
        for col in export_defs:
            key = str(col.get("key") or "")
            title = str(col.get("title") or key)
            display_title = title
            if display_title in used_titles:
                suffix = 2
                while f"{title}_{suffix}" in used_titles:
                    suffix += 1
                display_title = f"{title}_{suffix}"
            used_titles.add(display_title)
            out[display_title] = get_row_value(row, key)
        filtered.append(out)
    return filtered


def display_columns(export_defs: Sequence[Dict[str, Any]]) -> List[str]:
    result: List[str] = []
    for col in export_defs:
        title = str(col.get("title") or col.get("key") or "")
        if title in result:
            suffix = 2
            while f"{title}_{suffix}" in result:
                suffix += 1
            title = f"{title}_{suffix}"
        result.append(title)
    return result


def custom_hit_summary(rows: Sequence[Dict[str, Any]], export_defs: Sequence[Dict[str, Any]], custom_rules: Sequence[Dict[str, Any]]):
    title_by_key = {str(col.get("key")): str(col.get("title") or col.get("key")) for col in export_defs}
    items = []
    for rule in custom_rules:
        key = str(rule.get("name") or "")
        hit_count = sum(1 for row in rows if get_row_value(row, key) not in ("", None))
        items.append({"name": title_by_key.get(key, key), "key": key, "hit_count": hit_count, "row_count": len(rows)})
    return items


@app.get("/api/health")
def health():
    return {"status": "ok", "max_file_count": MAX_FILE_COUNT, "parse_concurrency": PARSE_CONCURRENCY}


@app.get("/api/columns")
def columns():
    system_columns = [
        "数量原始值",
        "单位",
        "数量",
        "Pos",
        "Material",
        "Price",
        "Amount",
        "Sales order ref",
        "Sales order ref num",
        "Sales order ref item",
        "包装描述",
        "催货标识",
        "DIM_CAR_BOX_INNER_LENGTH",
        "DIM_CAR_BOX_INNER_WIDTH",
        "DIM_CAR_BOX_INNER_HEIGHT",
        "长*宽*高",
        "轿门净高_HH",
        "轿厢净开门宽度_LL",
        "高*宽",
    ]
    groups = [
        {
            "name": "全部系统字段",
            "columns": system_columns,
        }
    ]
    default_keys = list(system_columns)
    return {
        "groups": groups,
        "default_columns": default_keys,
        "default_export_columns": [
            {"key": key, "title": key, "type": "system", "order": idx}
            for idx, key in enumerate(default_keys, start=1)
        ],
        "templates": {},
    }


@app.post("/api/preview")
async def preview(
    files: List[UploadFile] = File(...),
    columns: str = Form("[]"),
    custom_fields: str = Form("[]"),
    export_columns: str = Form("[]"),
):
    export_defs, custom_rules = parse_export_columns(export_columns, columns, custom_fields)
    if not export_defs:
        return JSONResponse(status_code=400, content={"message": "请至少添加一个导出列。"})

    temp_dir, pdf_paths, upload_errors = await save_uploads(files)
    try:
        if not pdf_paths:
            return JSONResponse(status_code=400, content={"message": "请上传有效 PDF 文件。", "errors": upload_errors})

        rows, parse_errors = await parse_all_files(pdf_paths, custom_rules)
        errors = upload_errors + parse_errors
        preview_rows = filter_rows(rows, export_defs)
        return {
            "columns": display_columns(export_defs),
            "export_columns": export_defs,
            "rows": preview_rows[:PREVIEW_ROW_LIMIT],
            "errors": errors,
            "custom_fields": custom_hit_summary(rows, export_defs, custom_rules),
            "summary": {
                "file_count": len(pdf_paths),
                "row_count": len(preview_rows),
                "shown_row_count": min(len(preview_rows), PREVIEW_ROW_LIMIT),
                "error_count": len(errors),
            },
        }
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@app.post("/api/export")
async def export_excel(
    files: List[UploadFile] = File(...),
    columns: str = Form("[]"),
    custom_fields: str = Form("[]"),
    export_columns: str = Form("[]"),
):
    export_defs, custom_rules = parse_export_columns(export_columns, columns, custom_fields)
    if not export_defs:
        return JSONResponse(status_code=400, content={"message": "请至少添加一个导出列。"})

    temp_dir, pdf_paths, upload_errors = await save_uploads(files)
    try:
        if not pdf_paths:
            shutil.rmtree(temp_dir, ignore_errors=True)
            return JSONResponse(status_code=400, content={"message": "请上传有效 PDF 文件。", "errors": upload_errors})

        rows, parse_errors = await parse_all_files(pdf_paths, custom_rules)
        errors = upload_errors + parse_errors

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = temp_dir / f"PDF_extract_result_{timestamp}.xlsx"
        await run_in_threadpool(write_excel, rows, export_defs, output_path, errors)

        return FileResponse(
            output_path,
            filename=output_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=BackgroundTask(shutil.rmtree, temp_dir, ignore_errors=True),
        )
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise


if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
else:
    @app.get("/")
    def root():
        return {"message": "API is running. Static frontend is missing."}
