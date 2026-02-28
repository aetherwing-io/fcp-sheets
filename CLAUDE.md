# fcp-sheets

## Project Overview
MCP server that lets LLMs create and edit spreadsheets through a semantic verb DSL.
Uses openpyxl as the native library (Tier 2 architecture).

## Architecture
- `src/fcp_sheets/model/` — Thin wrapper around openpyxl Workbook, cell ref parser, sheet index
- `src/fcp_sheets/server/` — Verb handlers (ops_*.py), queries, verb registry, resolvers
- `src/fcp_sheets/lib/` — Color palette, number formats, chart types, table styles
- `src/fcp_sheets/adapter.py` — FcpDomainAdapter bridging fcp-core to openpyxl
- `src/fcp_sheets/main.py` — Server entry point

## Key Patterns
- Each `ops_*.py` exports a `HANDLERS` dict mapping verb names to handler functions
- The adapter merges all HANDLERS at import time for dispatch
- `queries.py` exports `QUERY_HANDLERS` for query dispatch
- Block mode: `data` lines buffered in adapter, flushed on `data end`
- Undo/redo: byte snapshots via `wb.save(BytesIO)` / `load_workbook(BytesIO)`
- Batch atomicity: pre-batch snapshot, rollback on any op failure

## Commands
- `uv run pytest` — Run tests
- `uv run python -c "from fcp_sheets.main import main"` — Verify import
