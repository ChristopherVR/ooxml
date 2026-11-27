# Migrating to OOXML 2.5

OOXML 2.5 is a minor release: most applications upgrade from 2.1 by moving `ooxml-core`, `ooxml-ui` and the viewer packages they use to the versions released with it. This guide lists every change that can affect your code, configuration, tests or deployment, with examples of what to write in 2.5. See the [OOXML 2.5 release post](releases/2.5.md) for the new features.

The list comes from comparing the published 2.1.2 packages with this release: the names every entry point exports (resolved through its type declarations), the package manifests, the custom element tags, the events the elements dispatch and the `--office-*` tokens they declare.

- [What did not change](#what-did-not-change)
- [Upgrade the packages together](#upgrade-the-packages-together)
- [Excel command string tables moved behind `loadEditorStrings`](#excel-command-string-tables-moved-behind-loadeditorstrings)
- [Word view modes have six values](#word-view-modes-have-six-values)
- [The add-ins Store is off until the host enables it](#the-add-ins-store-is-off-until-the-host-enables-it)
- [Ribbon layouts follow Microsoft 365](#ribbon-layouts-follow-microsoft-365)
- [Ink brush widths are written in centimetres](#ink-brush-widths-are-written-in-centimetres)
- [More code loads on first use](#more-code-loads-on-first-use)
- [Commands disabled while a document is shared](#commands-disabled-while-a-document-is-shared)
- [Package size budgets](#package-size-budgets)

## What did not change

- **Entry points:** every subpath of `ooxml-core` and `ooxml-ui` in 2.1.2 still exists. `ooxml-core` adds `./docx/mail-merge` and `./automation/add-ins`.
- **Exports:** apart from the Excel string tables below, every name exported by 2.1.2 is still exported, from the same entry point.
- **Elements, events and tokens:** no custom element tag, element event or `--office-*` token was removed or renamed. PowerPoint adds two internal tags, `pptx-ui-contextual-ribbon` and `pptx-ui-picture-ribbon`, and `ooxml-ui` adds `office-ui-ink-canvas` and `office-ui-ink-pens` for the Draw tabs.
- **Bindings:** the React, Vue, Angular, Svelte, Solid and vanilla packages export the same components and props as before; Excel adds an optional `imageProvider` prop and Word an optional `onlineVideoHosts` prop.
- **Requirements:** Node.js 22.3 or newer, the last two major versions of Chrome, Edge, Firefox and Safari (Safari 16.4 at least), and the peer ranges of the bindings are unchanged.
- **Add-ins:** an add-in installs only after consent and runs only if you enable the runtime. Nothing an earlier host configured changes.
- **OpenTeams, Visio and the MCP servers:** nothing to change.

## Upgrade the packages together

The viewers are tested against the `ooxml-core` and `ooxml-ui` released with them, so move them in one change. `ooxml-ui` now pins `emf-converter` 5.0.1 (it pinned 5.0.0).

Before:

```json
{
	"dependencies": {
		"ooxml-core": "^2.1.2",
		"ooxml-ui": "^2.1.2"
	}
}
```

After:

```json
{
	"dependencies": {
		"ooxml-core": "^2.5",
		"ooxml-ui": "^2.5"
	}
}
```

`ooxml-core` no longer lists `xml-crypto` as an optional peer dependency: signature verification canonicalises XML itself. If you installed it only for `ooxml-core`, you can remove it.

Then check that one version of each is installed:

```sh
npm ls ooxml-core ooxml-ui emf-converter
```

## Excel command string tables moved behind `loadEditorStrings`

`ooxml-core/xlsx/ui` no longer exports the raw tables the Excel ribbon's translations were built from: `COMMAND_TABLES`, `commandStrings`, `CORE_FEATURE_STRINGS`, `DATA_TOOLS_STRINGS`, `FORMAT_CELLS_STRINGS`, `MESSAGE_STRINGS`, `NAVIGATION_STRINGS`, `RIBBON_HOME_STRINGS`, `RIBBON_TAB_STRINGS`, `RULES_STRINGS` and `TOOLS_STRINGS`. Each language is now built into its own file, and the French, German, Spanish and Simplified Chinese files load when that language is first used. `STRINGS[locale]` returns the English dictionary until `loadEditorStrings(locale)` has loaded the language, and `translate` falls back to English the same way.

Before:

```ts
import { commandStrings } from 'ooxml-core/xlsx/ui';

const labels = commandStrings('de');
console.log(labels['Insert Function']);
```

After:

```ts
import { loadEditorStrings, translate } from 'ooxml-core/xlsx/ui';

await loadEditorStrings('de');
console.log(translate('de', 'Insert Function'));
```

The Excel editor loads the language itself when you set `locale`; only code that reads the dictionaries directly needs this change. Keep `locale` at `en` to avoid the request.

## Word view modes have six values

Word now has Read Mode, Outline, Web Layout and Focus beside Draft and Print Layout. The `viewMode` of the Word navigation building blocks (`getNavigationProps()` and the bindings' building-block components) and the `onViewModeChange` callback take the new `DocumentView` type, `'draft' | 'print' | 'web' | 'read' | 'outline' | 'focus'`, instead of `'draft' | 'print'`. Code that switches on the value, or stores it, must accept the new views. `DOCUMENT_VIEWS` and `isDocumentView` are exported from `ooxml-ui/docx`.

Before:

```ts
const label = (mode: 'draft' | 'print') => (mode === 'print' ? 'Print Layout' : 'Draft');
const props = editor.getNavigationProps();
statusText.textContent = label(props.viewMode);
```

After:

```ts
import { DOCUMENT_VIEWS, type DocumentView } from 'ooxml-ui/docx';

const label = (mode: DocumentView) => DOCUMENT_VIEWS.find(([view]) => view === mode)?.[1] ?? mode;
const props = editor.getNavigationProps();
statusText.textContent = label(props.viewMode);
```

## The add-ins Store is off until the host enables it

Insert > Get Add-ins opens the shared Office Add-ins dialog in Word, Excel, PowerPoint and Visio. Its Store pivot is hidden, and Get Add-ins opens on My Add-ins, until the host enables a storefront for the editor or a container around it. No marketplace is contacted while the Store is off.

Before (2.1, no shared dialog):

```ts
// Nothing to configure: the products had no Get Add-ins dialog.
```

After:

```ts
import { setOfficeAddInStorefront } from 'ooxml-ui';

const editor = document.querySelector('xlsx-editor')!;
// Show the Store pivot with a Microsoft Marketplace handoff (opened only on a click),
setOfficeAddInStorefront(editor, { kind: 'microsoft' });
// or a catalogue of your own add-in files,
setOfficeAddInStorefront(editor, {
	kind: 'catalog',
	entries: [{ id: 'notes', label: 'Notes', read: () => fetchAddInFile('notes.xml') }],
});
// or turn it off again.
setOfficeAddInStorefront(editor, 'off');
```

Tests that expected a Store tab must enable it first; tests of the default should expect My Add-ins to be selected.

Three more add-in settings are new, and each is off or local until you choose otherwise:

- **Installed add-ins are remembered.** Upload My Add-in installs an XML manifest, a unified `manifest.json` or a `.zip` package after a consent step. An editor you embed keeps installs in memory for the session. The suite and the demos pass the IndexedDB store from `ooxml-ui/suite`, which keeps up to 16 MiB of add-in files for each profile on the device. Pass a store of your own with `setOfficeAddInPreferences(editor, store)`, or `null` to keep nothing.
- **Task panes run only when you turn the runtime on.** `setOfficeAddInRuntime(editor, { origins, frameSrcEnforced: true })` runs an installed add-in's pane in a sandboxed frame from the HTTPS origins you list. Set `frameSrcEnforced` only after your page's Content Security Policy restricts `frame-src` to those origins, redirects included; without it the runtime blocks every pane. `setOfficeAddInRuntime(editor, null)` turns execution off again.
- **Marketplace links resolve through your backend, or not at all.** `setOfficeAddInMarketplaceResolver` receives a Marketplace link's public asset id and returns the package; the library makes no request of its own. Pass `null` to turn resolution off.

If your page sets a Content Security Policy, add the publisher origins you approve to `frame-src` before you enable the runtime; nothing else in the policy changes.

## Ribbon layouts follow Microsoft 365

Several groups moved to where Microsoft 365 puts them. Every ribbon ID that existed in 2.1 still hides what it hid, so `hiddenTabs` and `hiddenGroups` settings keep working, but screenshots, browser tests and user documentation that name a group's position need updating.

- **PowerPoint Home:** the Arrange menu moved into the Drawing group as one vertical menu, and the Comments group left Home; the header **Comments** button now toggles a docked Comments pane. `hiddenGroups: ['home.arrange']` still hides the menu and its commands, `home.drawing.arrange` hides only the trigger, and `home.arrange.align-left` and the other row IDs hide single rows. Video, audio and SmartArt selections show the new Video Format, Audio Format, Playback and SmartArt Format tabs.
- **Excel Insert:** Add-ins sits after Illustrations, and Insert gains Shapes, Text Box, New Note, Header & Footer, PivotChart, Recommended PivotTables and Equation. PivotTable and Sparkline editing moved to their contextual tabs.
- **Word Insert:** the groups follow Word's order and fold in priority order at narrower widths, and the Text group gains Quick Parts and Signature Line. Table commands moved to the contextual Table Design and Table Layout tabs.
- **Word and Excel Draw:** a Draw tab sits after Insert; `hiddenTabs: ['draw']` hides it. Word shows a contextual Equation tab (`equation`) while an equation is selected.

Before:

```ts
// A browser test that found Arrange on PowerPoint's Home tab.
await page.getByRole('tab', { name: 'Home' }).click();
await page.getByRole('button', { name: 'Arrange' }).click();
```

After:

```ts
// Arrange is a menu in Home's Drawing group.
await page.getByRole('tab', { name: 'Home' }).click();
await page.locator('[data-ribbon-control="home.drawing.arrange"]').click();
```

## Ink brush widths are written in centimetres

Word and Excel now read, draw and save ink, on the InkML reader and writer PowerPoint already used. Saved ink states its brush width and height in centimetres, with InkML `transparency` beside the opacity, because Word, Excel and PowerPoint draw a width without a unit as a grey hairline. Ink written by 2.1 still reads back at its old width (a width without a unit is pixels), and saving it again writes centimetres. A 3 pixel pen:

Before:

```xml
<ink:brushProperty name="width" value="3" />
```

After:

```xml
<ink:brushProperty name="width" value="0.079375" units="cm" />
<ink:brushProperty name="height" value="0.079375" units="cm" />
```

There is nothing to change when you load and save ink through `ooxml-core`. Code of your own that reads these parts and assumed pixel widths must honour `units="cm"`.

## More code loads on first use

To keep the editors small, more code now loads from your own origin the first time it is used: Excel's non-English languages, chart tools, PivotTable dialogs and comments pane; Word's mail merge (including a recipient worker), References and table tools, Design theme data, Draw tab, Equation tools, Quick Parts and chart dialogs; Excel's Draw tab; the InkML part reader and writer the three Draw tabs share; PowerPoint's ink tools, slide previews and legacy `.ppt` writer; Visio's embedded chart and metafile code. Bundlers handle this without configuration. If you copy the built files by hand or serve them through an allowlist, deploy every file of the package's `dist` folder, keep their relative paths, and allow same-origin scripts and workers in your Content Security Policy.

Before:

```sh
# Copied only the entry files.
cp node_modules/ooxml-ui/dist/xlsx/index.js public/vendor/
```

After:

```sh
# Copy the whole build so on-demand chunks resolve beside the entries.
cp -r node_modules/ooxml-ui/dist public/vendor/ooxml-ui
cp -r node_modules/ooxml-core/dist public/vendor/ooxml-core
```

## Commands disabled while a document is shared

Some new data is not mapped to collaboration yet, so its commands are disabled, with the reason as their description, while a collaboration session is bound: Excel PivotTable editing and Place in Cell, and Word's Outline promote, demote and move commands (also disabled while changes are tracked). Ink drawn in Excel is not sent to the room; it stays in the window that drew it and is saved from there. Nothing changes for documents that are not shared.

There is no code to change. A host that shows its own controls for these commands can read the reason from the command's accessible description (`aria-description`).

## Package size budgets

The gzipped size budgets in `scripts/perf-budgets.json` moved with the new features. They limit what CI accepts; your bundle's real size depends on what you import. The largest moves since 2.1.2:

| Entry                                 | 2.1.2 budget | 2.5 budget | Why                                                    |
| ------------------------------------- | ------------ | ---------- | ------------------------------------------------------ |
| suite, opening a new Excel workbook   | 2,552 KiB    | 2,006 KiB  | languages, chart tools and PivotTable dialogs deferred |
| suite, opening a new Visio drawing    | 2,101 KiB    | 1,562 KiB  | embedded charts and metafiles deferred                 |
| `ooxml-core/visio`                    | 669 KiB      | 567 KiB    | the same                                               |
| `ooxml-core/pptx`                     | 1,494 KiB    | 1,470 KiB  | legacy writer and previews deferred                    |
| `ooxml-core/xlsx`                     | 831 KiB      | 890 KiB    | PivotCharts, rich values and equations                 |
| `ooxml-core/docx`                     | 574 KiB      | 601 KiB    | design, styles and arrangement                         |
| `ooxml-ui/docx`                       | 580 KiB      | 616 KiB    | Mailings, Design, References and views (most deferred) |
| `ooxml-core/automation/add-ins` (new) |              | 18 KiB     | add-in manifest validation                             |
| `ooxml-core/docx/mail-merge` (new)    |              | 43 KiB     | the mail merge engine                                  |

Nothing to change in your code; if your own build enforces size limits, measure it again.
