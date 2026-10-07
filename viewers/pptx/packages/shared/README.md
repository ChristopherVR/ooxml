# pptx-viewer-shared

This private package is a compatibility facade for `ooxml-ui/pptx`. It owns no
renderer, document operations, translations or browser components. All five
framework bindings import the public UI package directly.

The shared renderer and browser lifecycle live in `src/ui/src/pptx` at the
monorepo root. DOM-free editing and collaboration live in `src/core/pptx/editor`,
and assistant/MCP schemas live in `src/core/pptx/automation/schemas`.

## Public APIs

Use these published entries when sharing PowerPoint features with another project:

- `ooxml-ui/pptx`: renderer and web components.
- `ooxml-ui/pptx/theme`, `/loader`, `/i18n`, `/ai`: focused browser APIs.
- `ooxml-core/pptx/editor/<module>`: document operations.
- `ooxml-core/pptx/automation/schemas`: tool contracts.

The root, theme, loader, i18n and ai entries here only re-export those APIs.
Do not add implementations or new dependencies here. This facade is never
published to npm. See [the migration guide](../../../../docs/pptx-shared-migration.md)
for ownership, build order and provenance.

## Development

Build core and UI from the monorepo root before checking the facade:

```bash
bun run build
bun run --cwd src/ui build
bun run --cwd viewers/pptx/packages/shared build
bun run --cwd viewers/pptx/packages/shared typecheck
bun run --cwd viewers/pptx/packages/shared test
```

`test` and `test:watch` forward to the migrated product suite in `src/ui`.

## License

[Apache-2.0](LICENSE). Keep the [NOTICE](NOTICE) file with redistributions.
