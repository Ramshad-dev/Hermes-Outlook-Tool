from __future__ import annotations

import logging
from typing import Any

from .cli import register_cli
from .config import is_configured
from .schemas import TOOL_SCHEMAS
from .tools import TOOL_HANDLERS

logger = logging.getLogger(__name__)


def _tools_available() -> bool:
    return is_configured()


def register(ctx: Any) -> None:
    for name, schema in TOOL_SCHEMAS.items():
        handler = TOOL_HANDLERS.get(name)
        if handler is None:
            logger.error("Outlook tool %s has no handler; skipping", name)
            continue
        ctx.register_tool(
            name=name,
            toolset="outlook",
            schema=schema,
            handler=handler,
            description=schema.get("description", ""),
            check_fn=_tools_available,
        )
    register_cli(ctx)
