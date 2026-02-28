"""fcp-sheets — Spreadsheet File Context Protocol MCP server.

Custom server wiring (instead of create_fcp_server) for batch atomicity (C7).
"""

from __future__ import annotations

from fcp_core import (
    EventLog,
    VerbRegistry,
    format_result,
    parse_op,
    ParseError,
)
from fcp_core.session import SessionDispatcher
from fcp_core.server import _AdapterSessionHooks, _build_tool_description

from mcp.server.fastmcp import FastMCP
from mcp.types import TextContent

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SnapshotEvent
from fcp_sheets.server.reference_card import EXTRA_SECTIONS
from fcp_sheets.server.verb_registry import VERBS


def create_sheets_server() -> FastMCP:
    """Create the fcp-sheets MCP server with batch atomicity."""
    adapter = SheetsAdapter()

    registry = VerbRegistry()
    registry.register_many(VERBS)

    event_log: EventLog = EventLog()

    hooks = _AdapterSessionHooks(adapter)
    session = SessionDispatcher(
        hooks=hooks,
        event_log=event_log,
        reverse_event=adapter.reverse_event,
        replay_event=adapter.replay_event,
    )

    mcp = FastMCP(
        name="sheets-fcp",
        instructions="Spreadsheet File Context Protocol. Call sheets_help for the reference card.",
    )

    tool_description = _build_tool_description("sheets", registry, EXTRA_SECTIONS)
    reference_card = registry.generate_reference_card(EXTRA_SECTIONS)

    @mcp.tool(name="sheets", description=tool_description)
    def execute_ops(ops: list[str]) -> TextContent:
        if session.model is None:
            return TextContent(
                type="text",
                text=format_result(False, "No model loaded. Use sheets_session 'new' or 'open' first."),
            )

        # Take pre-batch snapshot for atomicity (C7)
        pre_batch = session.model.snapshot()
        results: list[str] = []

        for i, op_str in enumerate(ops):
            parsed = parse_op(op_str)
            if isinstance(parsed, ParseError):
                # Rollback on parse error
                session.model.restore(pre_batch)
                adapter.rebuild_indices(session.model)
                msg = (
                    f"! Batch failed at op {i + 1}: {op_str}. "
                    f"Error: {parsed.error}. "
                    f"State rolled back ({i} ops reverted)."
                )
                return TextContent(type="text", text=msg)

            result = adapter.dispatch_op(parsed, session.model, session.event_log)

            if not result.success:
                # Rollback on op failure (unless it's a data block accumulation)
                if result.message:  # Non-empty message = real error
                    session.model.restore(pre_batch)
                    adapter.rebuild_indices(session.model)
                    msg = (
                        f"! Batch failed at op {i + 1}: {op_str}. "
                        f"Error: {result.message}. "
                        f"State rolled back ({i} ops reverted)."
                    )
                    return TextContent(type="text", text=msg)

            formatted = format_result(result.success, result.message, result.prefix)
            if formatted.strip():  # Skip empty results from block accumulation
                results.append(formatted)

        return TextContent(type="text", text="\n".join(results))

    @mcp.tool(name="sheets_query")
    def execute_query(q: str) -> TextContent:
        """Query sheets state. Read-only."""
        if session.model is None:
            return TextContent(type="text", text=format_result(False, "No model loaded."))
        return TextContent(type="text", text=adapter.dispatch_query(q, session.model))

    @mcp.tool(name="sheets_session")
    def execute_session(action: str) -> TextContent:
        """sheets lifecycle: new, open, save, checkpoint, undo, redo."""
        return TextContent(type="text", text=session.dispatch(action))

    @mcp.tool(name="sheets_help")
    def get_help() -> TextContent:
        """Returns the sheets FCP reference card."""
        return TextContent(type="text", text=reference_card)

    return mcp


mcp = create_sheets_server()


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
