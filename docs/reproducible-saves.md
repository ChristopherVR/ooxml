# Reproducible saves and default theme fonts

Office stamps the time of a save and draws some ids at random. The core writers keep that
behaviour by default, and every writer takes the same options to make its output reproducible.

## The shared save options

`ReproducibleSaveOptions` (`ooxml-core/opc`, `src/core/opc/reproducible.ts`):

| Option    | What it fixes                                                                                                 | Default                                   |
| --------- | ------------------------------------------------------------------------------------------------------------- | ----------------------------------------- |
| `now`     | The time the save stamps: `dcterms:modified` (and a new package's `dcterms:created`), Visio's `TimeRefreshed` | the time of the save                      |
| `ids`     | The source of the ids Office draws at random (see below)                                                      | `randomIdSource` (Web Crypto)             |
| `zipDate` | The date of every ZIP entry                                                                                   | 1980-01-01 00:00:00 UTC, as Office writes |

`seededIdSource(seed)` gives the same sequence for the same seed. A source moves on with every id,
so give each save a fresh one. Two saves of the same input with the same `now` and a fresh
`seededIdSource(seed)` are byte-identical.

```ts
import { seededIdSource } from 'ooxml-core/opc';

const bytes = await saveDocx(model, media, {
	now: new Date('2026-01-15T09:00:00Z'),
	ids: seededIdSource(1),
});
```

Where each product takes them:

- Word: `saveDocx(model, media, options)` (`SaveOptions extends ReproducibleSaveOptions`). `now` also
  dates `modified` when `options.modified` asks for the stamp. Ids: content control `w:id`s made at
  save, the bibliography data store `ds:itemID`, glossary `w:guid`s, SmartArt model ids.
- PowerPoint: `PptxHandler.save(slides, options)` (`PptxHandlerSaveOptions extends ReproducibleSaveOptions`).
  Ids: table `a16:colId`, SmartArt model ids, chart series `c16:uniqueId`, modern comment author
  ids. `fixedDate` stays as a shorthand for all three (it also keeps `cp:revision` at 1).
  `PresentationOptions.now` dates a new deck's `docProps/core.xml`.
- Excel: `saveXlsx` stamps no time and draws no id at save; its entries are dated 1980-01-01.
- Visio: `editVsdx(bytes, edits, { now })` stamps recordsets with `now`; entries are always dated
  1980-01-01 (`zipDate` is not taken by Visio edits).
- Embedded chart workbooks (Word and PowerPoint) are dated with the host save's `zipDate`.

Inside a save, writers read the environment from the package they write
(`saveEnvironmentOf(zip)`, bound by `bindSaveEnvironment`), so no global state is shared between
concurrent saves.

## Default theme fonts

`OFFICE_THEME_FONTS` (`ooxml-core/drawingml`, `src/core/drawingml/office-theme-fonts.ts`) is the one
place a new document's theme fonts come from. New files name them only where Office's own new files
do and refer to them through the theme elsewhere:

| Product    | Where the face is named                                                  | Elsewhere                                                                                                | Override                                  |
| ---------- | ------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------- | ----------------------------------------- |
| Word       | `word/theme/theme1.xml` (a new document is saved with the default theme) | styles use `minorHAnsi` / `majorHAnsi`                                                                   | `model.theme` (`defaultWordTheme(fonts)`) |
| PowerPoint | `ppt/theme/theme1.xml`                                                   | masters use `+mj-lt` / `+mn-lt`; bullets follow the text font (`a:buFontTx`); SmartArt runs use `+mn-lt` | `PresentationOptions.theme.fonts`         |
| Excel      | the theme and the Normal style (`scheme="minor"`), as Excel does         |                                                                                                          | `createWorkbook({ themeFonts })`          |
| Charts     | the embedded workbook's Normal style, the host theme's body font         |                                                                                                          | follows the host document's theme         |
| Visio      | the document's face table (`FaceNames`)                                  |                                                                                                          | `createVsdx({ themeFonts: { minor } })`   |

No font is shipped: a reader without the face substitutes its own.

PowerPoint's own blank deck names Arial for the master's bullets; the core writes `a:buFontTx`
(the bullet follows the theme body font) so that a new deck names faces only in its theme.
