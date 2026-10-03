# ooxml-mcp

Combined MCP view of the format tools owned by the viewer repositories.
No schemas, document algorithms or mutations are duplicated here.

Install `ooxml-mcp` and the format packages you want to expose:
`docx-viewer-mcp`, `xlsx-viewer-mcp`, `visio-viewer-mcp` and
`pptx-viewer-mcp` (the version exposing `registerTools`).

Run `ooxml-tools /path/to/documents` for all four formats, or
`ooxml-tools /path/to/documents docx,xlsx,visio` for a selected set.
A missing selected package is a startup error, so tools are never silently omitted.
The new format tools require the upcoming core automation release (0.11.0+).

Each standalone package exports `registerTools(server, { rootDir })`.
The combined server loads those exact functions. Core alone owns parsing,
document models, editing, recalculation, preservation and serialization.
PowerPoint retains its existing unprefixed tool names. The combined server
passes the same document root to all four tool registrations.
Other tools use format prefixes (`docx_`, `xlsx_`, `visio_`).

Programmatic use: `await createServer({ rootDir, formats: ['docx', 'xlsx'] })`.
Importing this package does not connect stdio. All documents remain local.
This package ships JavaScript source and needs no separate build.
