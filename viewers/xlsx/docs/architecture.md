# Architecture and ownership

## Dependency direction

```text
published, one per framework (self-contained bundles)
  @christophervr/xlsx-react-viewer | xlsx-vue-viewer | xlsx-angular-viewer
  xlsx-svelte-viewer | xlsx-solid-viewer | xlsx-vanilla-viewer
  ----------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
               <xlsx-editor> web component (grid painter, ribbon, dialogs)
                         |
        ooxml-core (external, published)                 ooxml-ui (external)
   /xlsx       model, formula engine, layout, edit,      office-ui-* controls
               read/write .xlsx
   /xlsx/load  format detection, .xls (ole2 codecs inlined), CSV
        ^
   @christophervr/xlsx-core (external): thin re-export of ooxml-core/xlsx and /xlsx/load
```

Only `@christophervr/xlsx-core` and the six framework packages are published. `web-component` and `bindings` are `private` workspace packages: they are the shared source of the editor, never an npm install target. `scripts/build-packages.mjs` bundles them into each framework package, so a tarball imports only `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` and its framework peer. `scripts/check-published-refs.mjs` and the pack smoke test fail the build if an internal package or `ole2` leaks into a tarball.

## The core owns the logic

Everything that is not UI lives in the `xlsx` area of `ooxml-core` (source in the `ChristopherVR/ooxml` repository):

- the workbook model (`Workbook`, `Worksheet`, `Cell`, resolved and de-duplicated cell styles),
- addresses and ranges (`parseAddress`, `formatRange`, ...),
- number formats (`formatValue`, `parseCellInput`),
- the formula parser and calculation engine (`createCalcEngine`, dependency tracking, dynamic arrays),
- editing commands with undo and redo (`createEditSession`),
- DOM-free view logic: grid metrics, visible cells, cell views, conditional formats, chart view models,
- the `.xlsx` reader and writer, and, under `/xlsx/load`, format detection, the legacy `.xls` reader (through the shared ole2 codecs) and CSV.

A loaded workbook keeps the parts the model does not represent in `workbook.source`, so saving carries them through.

## The web component paints

`packages/web-component` is the only editing UI: it creates an edit session over the workbook, paints what the core's layout functions return, routes keyboard, mouse and ribbon input to editing commands, and owns styling, locales and dialogs. Shadow DOM scopes its styles; registration is browser-only and idempotent. Browser tests exercise the same contract through every binding.

`packages/bindings` owns framework integration only: mount and unmount, property updates, event forwarding and imperative handles. New editor features belong in the web component (or in the core); a binding change is justified only by a framework lifecycle or API requirement.

## Build and release boundary

`@christophervr/xlsx-core` stays a real dependency of every framework package rather than being bundled: applications import `createWorkbook` and the model types from it directly, so one shared copy keeps `Workbook` identical on both sides. `ooxml-core` and `ooxml-ui` are real dependencies too. Everything else internal is bundled; ole2 is inlined inside the core and is not a dependency of anything here. See [releasing](/releasing).
