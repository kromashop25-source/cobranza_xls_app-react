from __future__ import annotations

import logging
import os
import time
from uuid import uuid4

from fastapi import FastAPI, Request

logger = logging.getLogger("cobranza.observability")


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def _init_sentry() -> None:
    dsn = os.getenv("SENTRY_DSN", "").strip()
    if not dsn:
        return

    try:
        import sentry_sdk
        from sentry_sdk.integrations.fastapi import FastApiIntegration
        from sentry_sdk.integrations.logging import LoggingIntegration

        traces_sample_rate = float(os.getenv("SENTRY_TRACES_SAMPLE_RATE", "0.0"))
        profiles_sample_rate = float(os.getenv("SENTRY_PROFILES_SAMPLE_RATE", "0.0"))

        sentry_sdk.init(
            dsn=dsn,
            environment=os.getenv("APP_ENV", "production"),
            integrations=[
                FastApiIntegration(),
                LoggingIntegration(level=logging.INFO, event_level=logging.ERROR),
            ],
            traces_sample_rate=traces_sample_rate,
            profiles_sample_rate=profiles_sample_rate,
            send_default_pii=False,
        )
        logger.info("Sentry initialized.")
    except Exception as exc:
        logger.warning("Sentry setup skipped: %s", exc)


def _init_otel(app: FastAPI) -> None:
    if not _env_bool("OTEL_ENABLED", default=False):
        return

    try:
        from opentelemetry import trace
        from opentelemetry.exporter.otlp.proto.http.trace_exporter import OTLPSpanExporter
        from opentelemetry.instrumentation.fastapi import FastAPIInstrumentor
        from opentelemetry.sdk.resources import Resource
        from opentelemetry.sdk.trace import TracerProvider
        from opentelemetry.sdk.trace.export import BatchSpanProcessor, ConsoleSpanExporter

        service_name = os.getenv("OTEL_SERVICE_NAME", "cobranza-xls-app")
        resource = Resource.create({"service.name": service_name})
        tracer_provider = TracerProvider(resource=resource)

        endpoint = os.getenv("OTEL_EXPORTER_OTLP_ENDPOINT", "").strip()
        if endpoint:
            exporter = OTLPSpanExporter(endpoint=endpoint)
        else:
            exporter = ConsoleSpanExporter()

        tracer_provider.add_span_processor(BatchSpanProcessor(exporter))
        trace.set_tracer_provider(tracer_provider)

        FastAPIInstrumentor.instrument_app(app)
        logger.info("OpenTelemetry initialized.")
    except Exception as exc:
        logger.warning("OpenTelemetry setup skipped: %s", exc)


def _install_http_middleware(app: FastAPI) -> None:
    @app.middleware("http")
    async def with_request_context(request: Request, call_next):
        request_id = request.headers.get("x-request-id") or uuid4().hex[:12]
        started = time.perf_counter()

        try:
            response = await call_next(request)
        except Exception:
            elapsed_ms = (time.perf_counter() - started) * 1000
            logger.exception(
                "Unhandled error request_id=%s method=%s path=%s duration_ms=%.2f",
                request_id,
                request.method,
                request.url.path,
                elapsed_ms,
            )
            raise

        elapsed_ms = (time.perf_counter() - started) * 1000
        response.headers["x-request-id"] = request_id
        response.headers["x-process-time-ms"] = f"{elapsed_ms:.2f}"

        logger.info(
            "request_id=%s method=%s path=%s status=%s duration_ms=%.2f",
            request_id,
            request.method,
            request.url.path,
            response.status_code,
            elapsed_ms,
        )
        return response


def setup_observability(app: FastAPI) -> None:
    _init_sentry()
    _init_otel(app)
    _install_http_middleware(app)


def health_payload() -> dict[str, str | bool]:
    return {
        "ok": True,
        "sentry": bool(os.getenv("SENTRY_DSN", "").strip()),
        "otel": _env_bool("OTEL_ENABLED", default=False),
    }

