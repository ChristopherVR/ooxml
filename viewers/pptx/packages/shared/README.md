# pptx-viewer-shared

**A private compatibility facade that re-exports the public `ooxml-ui/pptx` APIs under their old package name. It is never published to npm and owns no code.**

The PowerPoint renderer, web components and browser lifecycle live in `src/ui/src/pptx` at the repository root (published as `ooxml-ui/pptx`), DOM-free editing in `src/core/pptx/editor` and the assistant and MCP tool schemas in `src/core/pptx/automation/schemas`. Its entries map onto them:

| Entry                       | Re-exports             |
| --------------------------- | ---------------------- |
| `pptx-viewer-shared`        | `ooxml-ui/pptx`        |
| `pptx-viewer-shared/theme`  | `ooxml-ui/pptx/theme`  |
| `pptx-viewer-shared/loader` | `ooxml-ui/pptx/loader` |
| `pptx-viewer-shared/i18n`   | `ooxml-ui/pptx/i18n`   |
| `pptx-viewer-shared/ai`     | `ooxml-ui/pptx/ai`     |

**Who imports it.** Only workspace code that is never published: some PowerPoint demos (`demos/pptx/demo-vanilla`, `demos/pptx/demo-host-review`) and the browser tests in `e2e/pptx`. The five framework bindings import `ooxml-ui` directly, and a published binding must not reference this package: `scripts/check-published-shared-refs.mjs` (`bun run check:published-types`) fails when a packed `.d.ts` imports from it. New code should import the `ooxml-ui` entries above; do not add implementations or dependencies here. See [the migration guide](../../../../docs/pptx-shared-migration.md) for ownership and build order.

## Development

Build the core and the UI at the repository root first:

```bash
bun run build
bun run --cwd src/ui build
bun run --cwd viewers/pptx/packages/shared build
bun run --cwd viewers/pptx/packages/shared typecheck
bun run --cwd viewers/pptx/packages/shared test # runs the PowerPoint suite in src/ui
```

## License

[Apache-2.0](LICENSE). Keep the [NOTICE](NOTICE) file with redistributions.
