from __future__ import annotations

import hashlib
from typing import Any, Iterable, Sequence

SOURCE_CACHE_SCHEMA = "source-cache-v1"


def source_artifact_key(
    *,
    name: str,
    data: bytes,
    mime: str = "",
    kind: str,
    extractor_version: str,
    options: dict[str, Any] | None = None,
) -> str:
    """Content-addressed key for one source-derived artifact.

    The key is based on exact source bytes, not filename/size alone. Callers keep the
    cache session-scoped; this helper deliberately contains no global/user state.
    """
    h = hashlib.sha256()
    for token in (SOURCE_CACHE_SCHEMA, kind, extractor_version, name or "", mime or ""):
        h.update(str(token).encode("utf-8", "ignore"))
        h.update(b"\0")
    h.update(hashlib.sha256(bytes(data or b"")).digest())
    h.update(b"\0")
    if options:
        for key in sorted(options):
            h.update(str(key).encode("utf-8", "ignore"))
            h.update(b"=")
            h.update(str(options[key]).encode("utf-8", "ignore"))
            h.update(b"\0")
    return h.hexdigest()


def bounded_cache_get(cache: dict[str, Any], key: str) -> Any:
    """LRU-like get for insertion-ordered dicts without altering cached content."""
    if key not in cache:
        return None
    value = cache.pop(key)
    cache[key] = value
    return value


def bounded_cache_set(cache: dict[str, Any], key: str, value: Any, *, max_entries: int) -> None:
    """Store exact value and evict only old cache entries, never source content."""
    if key in cache:
        cache.pop(key)
    cache[key] = value
    limit = max(1, int(max_entries or 1))
    while len(cache) > limit:
        oldest = next(iter(cache))
        cache.pop(oldest, None)


def ordered_parallel_results(items: Sequence[Any], fn, *, max_workers: int = 4) -> list[Any]:
    """Run independent deterministic preprocessing concurrently, preserving input order.

    This helper is intentionally generic and does not touch Streamlit/OpenAI state. The
    caller must only pass deterministic local preprocessing work here.
    """
    if len(items) <= 1 or max_workers <= 1:
        return [fn(item) for item in items]
    from concurrent.futures import ThreadPoolExecutor

    workers = min(max(1, int(max_workers)), len(items))
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="pa-source") as pool:
        futures = [pool.submit(fn, item) for item in items]
        return [future.result() for future in futures]
