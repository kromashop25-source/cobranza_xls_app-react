from __future__ import annotations

import os
from dataclasses import dataclass, asdict


@dataclass
class QueueStatus:
    mode: str
    redis_url_configured: bool
    arq_enabled: bool
    arq_installed: bool



def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}



def read_queue_status() -> QueueStatus:
    redis_url = os.getenv("REDIS_URL", "").strip()
    arq_enabled = _env_bool("ARQ_ENABLED", default=False)

    try:
        import arq  # noqa: F401

        arq_installed = True
    except Exception:
        arq_installed = False

    mode = "thread"
    if arq_enabled and redis_url and arq_installed:
        mode = "arq"

    return QueueStatus(
        mode=mode,
        redis_url_configured=bool(redis_url),
        arq_enabled=arq_enabled,
        arq_installed=arq_installed,
    )



def queue_status_payload() -> dict[str, str | bool]:
    return asdict(read_queue_status())
