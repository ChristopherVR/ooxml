# OOXML Office

The deployable Office application workspace (`ooxml-office`), shared by the integrated suite, standalone PWAs and desktop package. Core format logic remains in `ooxml-core`; reusable controls remain in `ooxml-ui`.

From the repository root:

```sh
bun install
bun run build:suite
bun run preview:suite
```

Output: `suite-dist/`. The root installs Office; `apps/word/`, `apps/excel/`, `apps/powerpoint/`, `apps/visio/` and `apps/teams/` each install independently. They share local profiles and documents when served from the same origin. Installation is available from the account menu or the browser's install command.

See [integration and limitations](../../docs/suite-integration.md) and [desktop packaging](../../desktop/README.md). This is a private application package, not an npm-published library. Cloud identity, cloud drive and document coauthoring are separate integration work.
