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
- [New host services are off until you supply them](#new-host-services-are-off-until-you-supply-them)
- [New files name their fonts through the theme](#new-files-name-their-fonts-through-the-theme)
- [Visio recordsets are numbered from 1](#visio-recordsets-are-numbered-from-1)
- [More is shared in collaboration](#more-is-shared-in-collaboration)
- [Package size budgets](#package-size-budgets)

## What did not change

- **Entry points:** every subpath of `ooxml-core` and `ooxml-ui` in 2.1.2 still exists. `ooxml-core` adds `./docx/mail-merge` and `./automation/add-ins`.
- **Exports:** apart from the Excel string tables below, every name exported by 2.1.2 is still exported, from the same entry point.
- **Elements, events and tokens:** no custom element tag, element event or `--office-*` token was removed or renamed. PowerPoint adds two internal tags, `pptx-ui-contextual-ribbon` and `pptx-ui-picture-ribbon`, and `ooxml-ui` adds `office-ui-ink-canvas` and `office-ui-ink-pens` for the Draw tabs.
- **Bindings:** the React, Vue, Angular, Svelte, Solid and vanilla packages export the same components and props as before; Excel adds an optional `imageProvider` prop and Word the optional `onlineVideoHosts`, `onlineVideoSearch`, `buildingBlocksStore` and `documentSigner` props.
- **Requirements:** Node.js 22.3 or newer, the last two major versions of Chrome, Edge, Firefox and Safari (Safari 16.4 at least), and the peer ranges of the bindings are unchanged.
- **Add-ins:** an add-in installs only after consent and runs only if you enable the runtime. Nothing an earlier host configured changes.
- **OpenTeams and the MCP servers:** nothing to change. Visio needs a change only if your code addresses an imported data recordset by its id (see below).

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

To keep the editors small, more code now loads from your own origin the first time it is used: Excel's non-English languages, chart tools, PivotTable dialogs and comments pane; Word's mail merge (including a recipient worker), References and table tools, Design theme data, Draw tab, Equation tools, Quick Parts and chart dialogs; Excel's Draw tab; the InkML part reader and writer the three Draw tabs share; PowerPoint's ink tools, slide previews, legacy `.ppt` writer and picture layouts; Visio's embedded chart and metafile code; the shared Ink Equation dialog, Excel's Picture Style 3-D drawing, and Word's label catalogue and Draw Table tool. Bundlers handle this without configuration. If you copy the built files by hand or serve them through an allowlist, deploy every file of the package's `dist` folder, keep their relative paths, and allow same-origin scripts and workers in your Content Security Policy.

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

## New host services are off until you supply them

Four Word features and Ink to Math in every product call services only your application can provide. Each is off, and its command disabled with the reason or hidden, until you set it:

- **Handwriting recognition:** `setInkMathRecognizer(recognizer)` from `ooxml-ui` (also exported by `ooxml-ui/pptx`) turns on Ink to Math in Word, Excel and PowerPoint and Ink Equation in Word. The recognizer receives only the strokes' X/Y points and returns LaTeX. PowerPoint's `configurePptxDraw({ recognizeMath })` overrides the shared one for that viewer.
- **Signing:** `documentSigner` takes a `PackageSigner` (a DER certificate and an RSA-SHA256 sign function); `signature.webCryptoPackageSigner` from `ooxml-core/opc` wraps a Web Crypto key. The library never reads a certificate store. Whether or not you set a signer, a document with a valid signature now opens read-only with an Edit Anyway bar, which warns that editing removes the signatures; tests that edit signed fixtures must click it first.
- **Online video search:** `onlineVideoSearch` takes a search provider. `createYouTubeVideoSearch({ apiKey })` from `ooxml-ui/docx` adapts the YouTube Data API with your own key; add the embed host it returns (`www.youtube-nocookie.com` by default) to `onlineVideoHosts`, and if your page sets a Content Security Policy, allow `connect-src https://www.googleapis.com` and the thumbnail host (`https://i.ytimg.com`) in `img-src`.
- **Shared building blocks:** `buildingBlocksStore` takes an object with `read()` and `write(value)`. `createBuildingBlocksStore(name)` from `ooxml-ui/suite` keeps them in an IndexedDB database of that name on the device, which the suite does for each profile.

Before:

```ts
// 2.1: no Ink to Math, signing, video search or shared building blocks to configure.
```

After:

```ts
import { setInkMathRecognizer } from 'ooxml-ui';
import { signature } from 'ooxml-core/opc';
import { createYouTubeVideoSearch } from 'ooxml-ui/docx';
import { createBuildingBlocksStore } from 'ooxml-ui/suite';

setInkMathRecognizer(recognize); // Word, Excel and PowerPoint

const editor = document.querySelector('docx-editor')!;
editor.documentSigner = signature.webCryptoPackageSigner(privateKey, certificateDer);
editor.onlineVideoHosts = ['www.youtube-nocookie.com'];
editor.onlineVideoSearch = createYouTubeVideoSearch({ apiKey });
editor.buildingBlocksStore = createBuildingBlocksStore('profile-1');
```

Each binding takes the same names as props. Set a value to `null` or `undefined`, or call `setInkMathRecognizer(undefined)`, to turn the service off again. `PRIVACY.md` describes what each one sends or keeps.

Mail merge also reads larger lists: local CSV and TSV recipient files are read in slices, and the editor's caps moved from 1,000 records, 64 fields and 5 MiB to 100,000 records, 255 fields and 32 MiB (20 MiB for `.xlsx`). Nothing changes for host data sources; tests that expected the old limits must use the new ones.

## New files name their fonts through the theme

A new Word document now saves a whole `word/settings.xml` with `compatibilityMode` 15, so Word opens it in its own mode rather than compatibility mode, and it is saved with the default Office theme. Its styles refer to the theme fonts, as Word 2013 and later write them, instead of naming Calibri. The default theme's fonts are still Calibri Light and Calibri; they come from `OFFICE_THEME_FONTS` in `ooxml-core/drawingml`, and new workbooks, Visio drawings and embedded chart workbooks take their fonts from the same place.

Before:

```xml
<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cs="Times New Roman"/>
```

After:

```xml
<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
```

Documents that came from a file keep their own settings, theme and styles. Code that reads a new document's fonts from `styles.xml` must resolve them through the theme. Tests that compare saved bytes can make saves reproducible with the shared save options, `{ now, ids: seededIdSource(seed) }` from `ooxml-core/opc` ([reproducible saves](reproducible-saves.md)); every ZIP entry is now dated 1980-01-01. To use other fonts in new files, pass `createWorkbook({ themeFonts })`, `createVsdx({ themeFonts })` or a Word theme of your own.

## Visio recordsets are numbered from 1

Visio's object model refuses a data recordset numbered 0, so a recordset imported with the `import-data-recordset` edit now takes the next id from 1, and its rowset writes the `_Visio_RowID_` column first (`c0`), as Visio saves it. The reader finds that column by name, so drawings saved by 2.1 still load with their ids.

Before:

```ts
// 2.1 gave the first recordset imported into a new drawing the id '0'.
const refresh = { type: 'refresh-data-recordset', pageId, recordsetId: '0', columns, rows };
```

After:

```ts
// 2.5 numbers it from '1', as Visio does; read the id from the drawing instead of assuming it.
const refresh = { type: 'refresh-data-recordset', pageId, recordsetId: '1', columns, rows };
```

## More is shared in collaboration

Excel now shares PivotTables, in-cell pictures and ink through its collaboration mapping, so the PivotTable commands and Place in Cell stay enabled while a workbook is shared, and ink drawn by one person reaches the others. Word's Outline promote, demote and move commands work while a document is shared and while changes are tracked. A few cases stay disabled with the reason as their description: a PivotTable over 4 Mi characters of JSON, a picture over 4 MB (`MAX_SHARED_CELL_PICTURE_BYTES`), Convert to Floating Picture, and an Outline level change that would need a new heading style.

Rooms are stamped when they carry the new data, so a room written by 2.1 never clears a joining peer's PivotTables or ink. A 2.1 client in the same room does not see them, though, so upgrade every client of a room together.

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
| `ooxml-ui/docx`                       | 580 KiB      | 643 KiB    | Mailings, Design, References and views (most deferred) |
| `ooxml-core/automation/add-ins` (new) |              | 18 KiB     | add-in manifest validation                             |
| `ooxml-core/docx/mail-merge` (new)    |              | 50 KiB     | the mail merge engine                                  |
| `ooxml-core/docx/pdf-export` (new)    |              | 36 KiB     | Word Save as PDF without the editor                    |
| `ooxml-core/pdf/write` (new)          |              | 27 KiB     | the shared vector PDF writer                           |

Nothing to change in your code; if your own build enforces size limits, measure it again.
