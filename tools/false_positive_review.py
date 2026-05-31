"""Post-process terminology QA issue sheets with an optional Codex false-positive review."""

from __future__ import annotations

import json
import re
import tempfile
from collections import OrderedDict
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from tools.codex_runner import run_codex_exec_prompt


OUTPUT_HEADERS = ("fp_decision", "fp_category", "fp_confidence", "fp_note", "fp_by")
VALID_DECISIONS = {"false_positive", "true_issue", "review"}
VALID_CONFIDENCE = {"high", "medium", "low"}
DEFAULT_CODEX_TIMEOUT_SECONDS = 600
DEFAULT_CODEX_CLUSTER_BATCH_SIZE = 40


@dataclass(frozen=True)
class ReviewColumnMapping:
    source_term_header: str
    expected_target_header: str
    issue_type_header: str
    source_text_header: str
    target_text_header: str
    eligible_issue_types: tuple[str, ...] = ()


@dataclass(frozen=True)
class ReviewExample:
    source_text: str
    target_text: str


@dataclass(frozen=True)
class ReviewCluster:
    key: str
    source_term: str
    expected_target: str
    issue_type: str
    source_text: str
    target_text: str
    row_numbers: tuple[int, ...]
    examples: tuple[ReviewExample, ...]


@dataclass(frozen=True)
class ReviewDecision:
    decision: str
    category: str
    confidence: str
    note: str


@dataclass(frozen=True)
class ReviewSummary:
    cluster_count: int
    reviewed_row_count: int
    failed_cluster_count: int = 0


Reviewer = Callable[[list[ReviewCluster]], dict[str, ReviewDecision]]
BatchReviewer = Callable[[list[ReviewCluster]], dict[str, ReviewDecision]]


TERM_PAIR_PROBLEM_MAPPING = ReviewColumnMapping(
    source_term_header="问题source术语",
    expected_target_header="预期target术语",
    issue_type_header="问题简述",
    source_text_header="source原文",
    target_text_header="target原文",
)

GLOSSARY_PROBLEM_MAPPING = ReviewColumnMapping(
    source_term_header="source术语",
    expected_target_header="期望target术语",
    issue_type_header="问题类型",
    source_text_header="source文本",
    target_text_header="target文本",
    eligible_issue_types=("术语未按术语表翻译",),
)


def cell_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def build_cluster_key(
    source_term: str,
    expected_target: str,
    issue_type: str,
    source_text: str,
    target_text: str,
) -> str:
    return json.dumps(
        [source_term, expected_target, issue_type, source_text, target_text],
        ensure_ascii=False,
        separators=(",", ":"),
    )


def get_header_indexes(worksheet) -> dict[str, int]:
    return {
        cell_text(worksheet.cell(row=1, column=column_index).value): column_index
        for column_index in range(1, worksheet.max_column + 1)
        if cell_text(worksheet.cell(row=1, column=column_index).value)
    }


def require_headers(header_indexes: dict[str, int], headers: Iterable[str]) -> None:
    missing_headers = [header for header in headers if header not in header_indexes]
    if missing_headers:
        raise ValueError(f"问题 sheet 缺少列: {'、'.join(missing_headers)}")


def ensure_output_columns(worksheet) -> dict[str, int]:
    header_indexes = get_header_indexes(worksheet)
    next_column = worksheet.max_column + 1
    output_columns: dict[str, int] = {}

    for header in OUTPUT_HEADERS:
        existing_column = header_indexes.get(header)
        if existing_column is not None:
            output_columns[header] = existing_column
            continue

        worksheet.cell(row=1, column=next_column, value=header)
        output_columns[header] = next_column
        next_column += 1

    return output_columns


def collect_review_clusters(
    worksheet,
    mapping: ReviewColumnMapping,
    *,
    sample_size: int = 5,
) -> list[ReviewCluster]:
    if sample_size < 1:
        raise ValueError("sample_size 必须大于等于 1。")

    header_indexes = get_header_indexes(worksheet)
    require_headers(
        header_indexes,
        (
            mapping.source_term_header,
            mapping.expected_target_header,
            mapping.issue_type_header,
            mapping.source_text_header,
            mapping.target_text_header,
        ),
    )

    grouped_rows: OrderedDict[str, list[tuple[int, ReviewExample, tuple[str, str, str, str, str]]]] = OrderedDict()
    for row_index in range(2, worksheet.max_row + 1):
        source_term = cell_text(worksheet.cell(row=row_index, column=header_indexes[mapping.source_term_header]).value)
        expected_target = cell_text(
            worksheet.cell(row=row_index, column=header_indexes[mapping.expected_target_header]).value
        )
        issue_type = cell_text(worksheet.cell(row=row_index, column=header_indexes[mapping.issue_type_header]).value)
        source_text = cell_text(worksheet.cell(row=row_index, column=header_indexes[mapping.source_text_header]).value)
        target_text = cell_text(worksheet.cell(row=row_index, column=header_indexes[mapping.target_text_header]).value)

        if mapping.eligible_issue_types and issue_type not in mapping.eligible_issue_types:
            continue
        if not source_term and not expected_target and not issue_type:
            continue

        key = build_cluster_key(source_term, expected_target, issue_type, source_text, target_text)
        grouped_rows.setdefault(key, []).append(
            (
                row_index,
                ReviewExample(source_text=source_text, target_text=target_text),
                (source_term, expected_target, issue_type, source_text, target_text),
            )
        )

    clusters: list[ReviewCluster] = []
    for key, rows in grouped_rows.items():
        source_term, expected_target, issue_type, source_text, target_text = rows[0][2]
        clusters.append(
            ReviewCluster(
                key=key,
                source_term=source_term,
                expected_target=expected_target,
                issue_type=issue_type,
                source_text=source_text,
                target_text=target_text,
                row_numbers=tuple(row_index for row_index, _, _ in rows),
                examples=tuple(example for _, example, _ in rows[:sample_size]),
            )
        )
    return clusters


def normalize_decision(raw_decision: dict[str, Any]) -> ReviewDecision:
    decision = cell_text(raw_decision.get("decision"))
    confidence = cell_text(raw_decision.get("confidence"))
    if decision not in VALID_DECISIONS:
        decision = "review"
    if confidence not in VALID_CONFIDENCE:
        confidence = "low"
    return ReviewDecision(
        decision=decision,
        category=cell_text(raw_decision.get("category")) or "需人工确认",
        confidence=confidence,
        note=cell_text(raw_decision.get("note")) or "Codex 未返回有效说明",
    )


def apply_false_positive_review_to_sheet(
    workbook,
    sheet_name: str,
    mapping: ReviewColumnMapping,
    *,
    reviewer: Reviewer,
    sample_size: int = 5,
) -> ReviewSummary:
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"找不到问题 sheet: {sheet_name}")

    worksheet = workbook[sheet_name]
    clusters = collect_review_clusters(worksheet, mapping, sample_size=sample_size)
    if not clusters:
        ensure_output_columns(worksheet)
        return ReviewSummary(cluster_count=0, reviewed_row_count=0)

    decisions = reviewer(clusters)
    output_columns = ensure_output_columns(worksheet)
    reviewed_row_count = 0
    failed_cluster_count = 0

    for cluster in clusters:
        decision = decisions.get(cluster.key)
        if decision is None:
            decision = ReviewDecision(
                decision="review",
                category="需人工确认",
                confidence="low",
                note="Codex 未返回该 cluster 的判定",
            )
            failed_cluster_count += 1

        for row_index in cluster.row_numbers:
            worksheet.cell(row=row_index, column=output_columns["fp_decision"], value=decision.decision)
            worksheet.cell(row=row_index, column=output_columns["fp_category"], value=decision.category)
            worksheet.cell(row=row_index, column=output_columns["fp_confidence"], value=decision.confidence)
            worksheet.cell(row=row_index, column=output_columns["fp_note"], value=decision.note)
            worksheet.cell(row=row_index, column=output_columns["fp_by"], value="codex")
            reviewed_row_count += 1

    return ReviewSummary(
        cluster_count=len(clusters),
        reviewed_row_count=reviewed_row_count,
        failed_cluster_count=failed_cluster_count,
    )


def build_codex_prompt(clusters: list[ReviewCluster], *, prompt_ids: list[str] | None = None) -> str:
    if prompt_ids is None:
        prompt_ids = [cluster.key for cluster in clusters]
    if len(prompt_ids) != len(clusters):
        raise ValueError("prompt_ids 数量必须和 clusters 一致。")

    cluster_payload = [
        {
            "id": prompt_id,
            "source_term": cluster.source_term,
            "expected_target": cluster.expected_target,
            "issue_type": cluster.issue_type,
            "source_text": cluster.source_text,
            "target_text": cluster.target_text,
            "examples": [
                {
                    "source_text": example.source_text,
                    "target_text": example.target_text,
                }
                for example in cluster.examples
            ],
        }
        for cluster, prompt_id in zip(clusters, prompt_ids, strict=True)
    ]
    return (
        "你是本地化 QA 术语误报筛查助手。\n\n"
        "任务：判断术语检查结果是否为 false positive。术语检查器只做字符串匹配，可能因为单复数、阴阳性、"
        "大小写、变音符号、标记包裹、词性变化、动名词/名词化、形容词化、法语句法重组，"
        "或 source 非术语用法而误报。\n\n"
        "核心原则：\n"
        "1. 这是术语一致性 QA，不是一般翻译质量评估。\n"
        "2. expected_target 是术语表要求的目标术语。\n"
        "3. 如果 source_text 中的 source_term 是明确专有术语、技能名、道具名、系统名、功能名、UI 固定名，"
        "target_text 应使用 expected_target 或其明确形态/格式变体。\n"
        "4. 但如果 source_term 在 source_text 中不是专有术语名，而是普通动作、状态、属性、修饰语、语法成分，"
        "target_text 可以使用 expected_target 的自然词性变化、动名词/名词化、形容词化、词族派生或法语句法重组。\n\n"
        "请只基于给定数据判断，不要改译文，不要新增术语，不要运行命令，不要读取文件。\n\n"
        "输入包含多个 clusters。每个 cluster 含 id、source_term、expected_target、issue_type、source_text、target_text、examples。"
        "source_text 和 target_text 是本次判断的主要依据；examples 是同一 cluster 的重复样本。\n\n"
        "请原样返回每个 cluster 的短 id。\n\n"
        "输出严格 JSON，不要 Markdown，不要解释文字：\n"
        "{\n"
        '  "results": [\n'
        "    {\n"
        '      "id": "原 cluster id",\n'
        '      "decision": "false_positive | true_issue | review",\n'
        '      "category": "单复数/阴阳性变体 | 词性变化/动名词/自然句法重组 | 格式/大小写/标记导致漏命中 | source非术语用法 | 同义译名/定稿差异但未按术语表 | 可能真术语问题 | 需人工确认",\n'
        '      "confidence": "high | medium | low",\n'
        '      "note": "一句中文说明"\n'
        "    }\n"
        "  ]\n"
        "}\n\n"
        "判断原则：\n"
        "1. 如果 target_text 已使用 expected_target 的明显单复数、阴阳性、大小写、变音符号差异、标记包裹形式，判 false_positive。\n"
        "2. 如果 target_text 没有逐字出现 expected_target，但使用了 expected_target 的明确词性变化、动名词、名词化、形容词化、词族派生或法语自然句法重组，并且含义仍对应 source_term，判 false_positive 或 review。\n"
        "   例：术语表是名词，但译文因句法需要使用对应动词/动作名词/形容词形式，可判 false_positive。\n"
        "   例：expected_target 的核心词根仍可识别，只是为了法语语法发生性数、词性或结构变化，可判 false_positive。\n"
        "3. 如果 source_term 在 source_text 中不是专有术语名，而是普通词、动词、形容词、状态描述或更大短语的一部分，且 target_text 用自然译法表达了该普通含义，判 false_positive 或 review。\n"
        "4. 如果 source_text 中的 source_term 是明确专有术语、技能名、道具名、系统名、功能名、UI 固定名，而 target_text 使用的是另一个同义译名、近义译名、自然改写、短称、看似官方的另一定稿或语义等价译名，但没有使用 expected_target 或其明确形态/格式变体，判 true_issue。\n"
        "   例：source_term=Fiery Assault，expected_target=Assaut Enflammé，target_text 使用 Assaut flamboyant；即使含义接近，也因为未按术语表使用 Assaut Enflammé，判 true_issue。\n"
        "5. 不要仅因为“含义一致”“语义等价”“译文自然”就判 false_positive。必须能说明它是 expected_target 的形态/词性/句法变体，或 source_term 在该处不是专有术语用法。\n"
        "6. 如果 target_text 完全没有表达 source_term 的术语含义，或用了明显不同概念，判 true_issue。\n"
        "7. 如果需要项目术语规范才能确认某个短称、另一定稿、词族派生或上下文用法是否允许，判 review。\n"
        "8. 不确定时优先 review，不要硬判 false_positive。\n\n"
        "现在请判断以下 clusters：\n"
        f"{json.dumps(cluster_payload, ensure_ascii=False, indent=2)}\n"
    )


def extract_json_object(text: str) -> dict[str, Any]:
    stripped = text.strip()
    fenced_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", stripped, flags=re.DOTALL)
    if fenced_match:
        stripped = fenced_match.group(1).strip()
    if not stripped.startswith("{"):
        start = stripped.find("{")
        end = stripped.rfind("}")
        if start == -1 or end == -1 or end < start:
            raise ValueError("Codex 输出中没有 JSON 对象。")
        stripped = stripped[start : end + 1]
    parsed = json.loads(stripped, strict=False)
    if not isinstance(parsed, dict):
        raise ValueError("Codex 输出 JSON 不是对象。")
    return parsed


def parse_codex_decisions(text: str) -> dict[str, ReviewDecision]:
    parsed = extract_json_object(text)
    results = parsed.get("results")
    if not isinstance(results, list):
        raise ValueError("Codex 输出缺少 results 数组。")

    decisions: dict[str, ReviewDecision] = {}
    for item in results:
        if not isinstance(item, dict):
            continue
        cluster_id = cell_text(item.get("id"))
        if not cluster_id:
            continue
        decisions[cluster_id] = normalize_decision(item)
    return decisions


def review_clusters_in_batches(
    clusters: list[ReviewCluster],
    *,
    batch_size: int,
    batch_reviewer: BatchReviewer,
    retry_missing: bool = False,
) -> dict[str, ReviewDecision]:
    if batch_size < 1:
        raise ValueError("batch_size 必须大于等于 1。")

    decisions: dict[str, ReviewDecision] = {}
    for start_index in range(0, len(clusters), batch_size):
        batch = clusters[start_index : start_index + batch_size]
        batch_decisions = batch_reviewer(batch)
        if retry_missing:
            missing_clusters = [cluster for cluster in batch if cluster.key not in batch_decisions]
            for cluster in missing_clusters:
                batch_decisions.update(batch_reviewer([cluster]))
        decisions.update(batch_decisions)
    return decisions


def review_clusters_with_codex(
    clusters: list[ReviewCluster],
    *,
    codex_command: str = "codex",
    model: str | None = None,
    reasoning_effort: str = "high",
    timeout_seconds: int = DEFAULT_CODEX_TIMEOUT_SECONDS,
    batch_size: int = DEFAULT_CODEX_CLUSTER_BATCH_SIZE,
) -> dict[str, ReviewDecision]:
    if not clusters:
        return {}

    return review_clusters_in_batches(
        clusters,
        batch_size=batch_size,
        batch_reviewer=lambda batch: review_cluster_batch_with_codex(
            batch,
            codex_command=codex_command,
            model=model,
            reasoning_effort=reasoning_effort,
            timeout_seconds=timeout_seconds,
        ),
        retry_missing=True,
    )


def review_cluster_batch_with_codex(
    clusters: list[ReviewCluster],
    *,
    codex_command: str = "codex",
    model: str | None = None,
    reasoning_effort: str = "high",
    timeout_seconds: int = DEFAULT_CODEX_TIMEOUT_SECONDS,
) -> dict[str, ReviewDecision]:
    prompt_id_to_cluster_key = {f"cluster-{index + 1}": cluster.key for index, cluster in enumerate(clusters)}
    prompt = build_codex_prompt(clusters, prompt_ids=list(prompt_id_to_cluster_key))
    with tempfile.TemporaryDirectory(prefix="tag-exactor-fp-") as tmp_dir:
        output_path = Path(tmp_dir) / "codex-output.txt"
        output_text = run_codex_exec_prompt(
            prompt,
            output_path=output_path,
            model=model,
            reasoning_effort=reasoning_effort,
            timeout_seconds=timeout_seconds,
            codex_command=codex_command,
            error_prefix="Codex 假阳性复核失败",
        )
    prompt_decisions = parse_codex_decisions(output_text)
    return {
        cluster_key: decision
        for prompt_id, decision in prompt_decisions.items()
        if (cluster_key := prompt_id_to_cluster_key.get(prompt_id)) is not None
    }
