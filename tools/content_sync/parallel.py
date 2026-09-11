"""Bounded file tasks; results and exceptions are consumed in input order."""
from collections import deque
from concurrent.futures import ThreadPoolExecutor


def validate_workers(workers: int) -> None:
    if isinstance(workers, bool) or not isinstance(workers, int) or not 1 <= workers <= 4:
        raise ValueError("并行任务数必须是 1–4 的整数。")


def _capture(function, item):
    try:
        return function(item), None
    except Exception as exc:
        # Do not retain worker tracebacks, workbook frames or parser streams.
        return None, str(exc)


def ordered_results(function, items, workers):
    """Keep at most workers files in flight, without sharing counters or callbacks."""
    validate_workers(workers)
    if workers == 1:
        for item in items:
            yield _capture(function, item)
        return
    iterator = iter(items)
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="content-sync") as pool:
        pending = deque()
        for _ in range(workers):
            item = next(iterator, None)
            if item is not None:
                pending.append(pool.submit(_capture, function, item))
        while pending:
            result = pending.popleft().result()
            item = next(iterator, None)
            if item is not None:
                pending.append(pool.submit(_capture, function, item))
            yield result
