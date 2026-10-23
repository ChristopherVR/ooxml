# Chart elements (`src/ui/src/chart`)

The chart editing UI every product mounts: the Format pane, the Edit Data sheet and the Chart
Design galleries. The elements are DOM only and hold no document logic. A product maps their
events onto its own chart model, and the neutral chart operations live in `src/core/chart`
(`chart/edit` for the typed intents and the style, palette and quick-layout presets).

| Element or export                                     | What it does                                                                                                                       |
| ----------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `<office-ui-chart-format-pane>`                       | The Format task pane (Chart Options and Text Options, element picker, collapsible sections). Accepts a `source` or a `chartSpace`. |
| `<office-ui-chart-data-sheet>`                        | The Edit Data table: categories down the first column, series across the header, paste from spreadsheets.                          |
| `<office-ui-chart-picker>` (`dialog/chart-picker.ts`) | The body of Insert Chart and Change Chart Type. It is the picker the product's dialog wraps; the dialog owns OK and Cancel.        |
| Chart Design galleries (`design-galleries.ts`)        | Chart Styles, Change Colors and Quick Layout as `office-ui-gallery` sections.                                                      |

## Format pane

Set `source` (a `ChartFormatSource`: the element targets, element views, series, axis, legend and
label options) or `chartSpace` (an `ooxml-core` `ChartSpace`, read through `chartSpaceFormatSource`),
then `target` (for example `chartArea`, `series:0`, `axis:1`). `theme` supplies the colour theme for
the colour choices. The pane applies nothing itself. Each committed field fires
`office-chart-format-edit` with `{ target, field, value, intent }`: `intent` is the typed
`ChartEditIntent` from `chartFormatFieldIntent`, or `null` for a field no neutral operation covers.
A host applies the intent (`applyChartEditIntent` runs it over a `ChartSpace`) or maps it onto its
own model, then sets `source` again, and the pane repaints from it. The pane also fires
`office-chart-format-target` when the element picker changes, and `office-chart-format-close` from
its Close button and Escape. `disabled` makes every control read-only.

The `source.slots.fill` entry leaves the fill section to the host: the host's light-DOM child with
that `slot` name is shown in its place. `source.picker` and `source.textOptions` change the element
picker's caption and whether Text Options is offered.

Captions are the `chartFormat.*` keys. `locale` (`en`, `fr`, `de`, `es`, `zh-CN`, or any BCP 47 tag
that maps onto one of them) selects the pane's own dictionary, and `translateLabel(key, params,
fallback)` overrides the captions a product words differently; it wins wherever it knows a key. The
dictionary is also exported on its own, so a product can build its dictionary on it, from
`ooxml-ui/chart-format-locales` (`CHART_FORMAT_STRINGS`, `CHART_FORMAT_LOCALES`, source in
`format-pane-locales/`).

## Data sheet

`data` is `{ categories, series: [{ name, values }] }`. The sheet shows edits at once and takes the
next `data` the host sets. Each change fires `office-chart-data-change` with `{ edits, data }`, where
`edits` are the neutral `ChartDataSheetEdit` operations (`applyChartDataSheetEdits` applies them to
the data). A product that holds its own model subclasses the element and overrides `emitEdits` to map
the edits onto that model as well. `highlight` marks a series (and optionally a point), `locale`
sets the number convention (the browser's when unset) and `disabled` makes it read-only. A value that
is not a number is refused and the cell restored.

## Chart Design galleries

`chartStyleGallerySection`, `chartColorGallerySections` and `chartQuickLayoutGallerySection` return
DOM-free section data with English labels and SVG tiles (from `gallery-tiles.ts`). A product maps
them onto its gallery control, translates the label templates, and applies a pick through
`applyChartStylePreset`, `applyChartColorPalette` or `applyChartQuickLayout` (in `src/core/chart`), or
its own mapping of those operations. The tile size constants (`CHART_STYLE_TILE`, `CHART_COLOR_TILE`,
`CHART_LAYOUT_TILE`) are what the galleries use.

## Files

- `format-pane*.ts`: the element (`format-pane.ts`), its model and view types, the field renderers
  (`format-pane-fields.ts`, `-fill-line.ts`, `-axis.ts`, `-effects.ts`, `-text.ts`, `-3d.ts`,
  `-space.ts`), the typed intents (`format-pane-intent.ts`) and the translations.
- `data-sheet*.ts`: the element, its keyboard model (`data-sheet-keys.ts`) and labels.
- `design-galleries.ts`, `gallery-tiles.ts`: the Chart Design gallery sections and tiles.

Each element is a Lit class on `OfficeElement` with its CSS next to it, as described in
`src/ui/README.md` ("How the elements are built"). The PowerPoint wrappers (`pptx-ui-chart-format-pane`,
`pptx-ui-chart-data-sheet`, `pptx-ui-insert-chart-dialog`) are in `src/ui/src/pptx/web-components`.
