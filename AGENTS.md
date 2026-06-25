# Repository Guidelines

## Project Structure & Module Organization

This is a small Python MCP server package. Core implementation lives in
`ms_teams_mcp/server.py`, with the package marker in `ms_teams_mcp/__init__.py`.
Tests are in `tests/test_server.py` and currently focus on tool formatting,
error handling, retry behavior, and Graph API call mocking. Project metadata and
the console entry point are defined in `pyproject.toml`. Design notes and
planning documents live under `docs/superpowers/`.

## Build, Test, and Development Commands

- `pip install -e .` installs the package locally with runtime dependencies.
- `pip install -e . pytest` installs the package plus the test runner.
- `pytest` runs the test suite; tests mock Microsoft Graph calls and should not
  require network access.
- `ms-teams-mcp` runs the MCP server over stdio.
- `ms-teams-mcp serve` runs the streamable HTTP transport on the default port.
- `ms-teams-mcp serve --transport sse` runs the SSE transport.
- `ms-teams-mcp auth` starts Device Code Flow authentication from the CLI.

## Coding Style & Naming Conventions

Use Python 3.10+ and keep source comments, docstrings, and user-facing strings in
English. The project has no configured formatter or linter, so match the existing
style in `server.py`: 4-space indentation, snake_case functions, uppercase module
constants, and concise helper functions prefixed with `_` for internal behavior.
All Microsoft Graph access should go through the shared Graph helpers
(`graph_get`, `graph_post`, `graph_patch`, `graph_delete`, etc.) so retry and
error handling remain consistent.

## Testing Guidelines

Use `pytest` and `unittest.mock` for new tests. Add tests to `tests/test_server.py`
unless a new module justifies a separate test file. Name tests by behavior, for
example `test_retries_on_429_then_succeeds`. Mock Graph and HTTP requests; do not
write tests that depend on live Microsoft credentials, token files, or network
state.

## Commit & Pull Request Guidelines

Recent history uses conventional prefixes such as `feat(teams): ...`,
`feat(briefing): ...`, `test(briefing): ...`, `docs(briefing): ...`, and
`chore(teams): ...`. Follow that pattern and keep subjects imperative and
specific. Pull requests should describe the user-visible change, note any new
Graph permissions or configuration changes, link related issues, and include the
test command run, typically `pytest`.

## Security & Configuration Tips

Required credentials are `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, and `MS_TENANT_ID`.
Never commit secrets or token cache files; local tokens are stored at
`~/.ms_mcp_token.json`. For tools that send, reply, forward, create, update, or
delete data, preserve the repository rule that users must explicitly confirm the
full content before the tool is called.
