"""UI-independent, optional operation diagnostics; never write report files."""

import json
from itertools import islice


def _brief(value, depth=0):
    """Bound diagnostics before JSON encoding, including huge unmatched-row lists."""
    if isinstance(value, str):
        return value if len(value) <= 1000 else value[:1000] + "…（截断）"
    if depth >= 4:
        return "…"
    if isinstance(value, dict):
        result = {key: _brief(item, depth + 1) for key, item in islice(value.items(), 20)}
        if len(value) > 20:
            result["省略字段数"] = len(value) - 20
        return result
    if isinstance(value, (list, tuple)):
        result = [_brief(item, depth + 1) for item in value[:12]]
        if len(value) > 12:
            result.append(f"…另有 {len(value) - 12} 项")
        return result
    return value


def emit_log(callback, message: str, details=None) -> None:
    if callback is not None:
        if details:
            message += " | " + json.dumps(_brief(details), ensure_ascii=False, default=str)
        callback(message)


def log_result(callback, result) -> None:
    """One compact record per file; paths shared by the batch are logged once."""
    if callback is None:
        return
    labels = {"updated": "已更新", "unchanged": "无变化", "read": "已读取",
              "replaced": "已替换", "copied": "已复制", "skipped": "跳过", "failed": "失败"}
    message = getattr(result, "message", "") or getattr(result, "error", "") or ""
    parts = [f"[{labels.get(result.status, result.status)}] {result.source}"
             + (f" — {message}" if message else "")]
    postprocess_error = getattr(result, "postprocess_error", None)
    if postprocess_error:
        parts.append(f"[重存失败] {postprocess_error}")
    details = getattr(result, "details", None) or vars(result)
    if "matched_rows" in details and result.status in {"updated", "unchanged"}:
        parts.append(f"匹配 {details['matched_rows']} 行，更新 {details['updated_cells']} 格")
        skipped = details.get("blank_source_cells", 0) + details.get("occupied_cells", 0)
        if skipped:
            parts.append(f"策略跳过 {skipped} 格")
        if details.get("unmatched_rows"):
            rows = details["unmatched_rows"]
            parts.append(f"未匹配 {len(rows)} 行（{', '.join(map(str, rows[:12]))}"
                         + ("…" if len(rows) > 12 else "") + "）")
        if details.get("invalid_rows"):
            parts.append(f"身份为空 {details['invalid_rows']} 行")
    elif "counts" in details:
        units, rows, total_units, total_rows = details["counts"]
        parts.append(f"未翻译量: {units}/{total_units}；未翻译行: {rows}/{total_rows}")
    elif "identities" in details:
        parts.append(f"有效身份 {details['identities']}")
    if details.get("backup"):
        parts.append("已备份")
    # Ambiguous file replacements still need their candidate paths for resolution.
    if "sources" in details:
        parts.append("重名候选 " + json.dumps(_brief(details), ensure_ascii=False, default=str))
    callback(" · ".join(parts).replace("\r", "\\r").replace("\n", "\\n"))


def log_summary(callback, summary, extra: str = "") -> None:
    emit_log(callback, f"完成：成功 {summary.succeeded_files}，失败 {summary.failed_files}，跳过 {summary.skipped_files}"
             + (f" · {extra}" if extra else ""))
    for warning in summary.warnings:
        emit_log(callback, f"[警告] {warning}")
