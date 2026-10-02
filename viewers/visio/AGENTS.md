# Agent guidance

This is the public, UI-only Visio beta viewer. GitHub Pages deployment and the initial npm placeholder packages are authorized. Functional npm releases require a released core dependency first. Full Microsoft Visio parity is a target, not a current claim.

- All VSDX parsing, packaging, ShapeSheet interpretation and document models live in `ooxml-core/visio` in the sibling `ooxml` repository. Do not fork that logic here.
- One SVG renderer and one custom element serve every framework. Bindings own lifecycle and prop/event forwarding only.
- Strict TypeScript, including exact optional properties and unchecked indexed access. Keep modules small and test parsing separately in the core repository.
- No uploads, telemetry or remote document fetches. Documents remain local to the browser. Never insert document XML/HTML into the DOM.
- Run `npm run check` and `npm run test:browser` before declaring completion. Record anything not tested. Unsupported features must produce honest diagnostics.
- Follow conventional commits. No em dashes in source, docs or UI. Do not commit generated output, credentials or user's documents.
- `docs/parity.md` is the evidence-based capability ledger. Edit the canonical table in `docs/parity.html`, then run `npm run docs:sync`. Documentation checks reject drift between the two copies. Generated fixture tests do not establish visual parity with Microsoft Visio.
