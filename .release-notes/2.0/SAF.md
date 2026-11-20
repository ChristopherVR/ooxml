## Added

- `ooxml-ui/pptx`: `installTouchPress(root)` turns a touch long-press into a `contextmenu` event and a double-tap into a `dblclick` event on controls that opt in with `data-pptx-touch-press="menu"`, `"dblclick"` or both. `createTouchPressRecognizer` is the pointer-event recogniser behind it.
- `ooxml-ui/pptx`: `copyPngToClipboard` writes a PNG to the clipboard within the click that asked for it, so Safari accepts the write.
- `ooxml-ui/pptx`: `printHtmlDocument` prints one HTML document from a hidden iframe, or from a print window on iOS.
- `ooxml-ui/pptx`: `pickVideoRecordingFormat` and `requireVideoRecordingFormat` choose the recorded video container (WebM, or MP4 where only MP4 records) with `MediaRecorder.isTypeSupported`. A browser that records neither gets a `VideoExportUnsupportedError`.
- `ooxml-ui/pptx`: `maxCanvasAreaFor` and `releaseCanvas` expose the canvas area cap the exporters honour (4096 x 4096 on iOS WebKit) and free a canvas's memory early.
- `docs/browser-support.md` records the browser floor (the last two major versions of Chrome, Edge, Firefox and Safari, Safari 16.4 at least) and the Safari behaviours the packages handle.

## Changed

- PowerPoint viewers (React, Vue, Angular, Svelte, vanilla): the video export is labelled Video instead of WebM in the ribbon, the export dialog and the File page, in every locale. The file is MP4 on Safari and WebM elsewhere.

## Fixed

- PowerPoint viewers (React, Vue, Angular, Svelte, vanilla): Copy Slide as Image works in Safari. Before, it failed because the clipboard write came after the slide render.
- PowerPoint viewers: downloads on iPadOS no longer fail. The download link stays valid until the user confirms the save.
- PowerPoint viewers: copying and pasting slides works in Safari, which has no custom clipboard formats. The slide package is also kept in the page, so slides move between viewers on the same page.
- PowerPoint viewers: large PNG and JPEG exports are no longer blank on iOS and iPadOS. Exports stay under WebKit's canvas area and memory limits there.
- PowerPoint viewers: Safari exports include pictures and embedded fonts again. The export waits for the images to decode and redraws once on WebKit.
- PowerPoint viewers: video export records MP4 on Safari instead of failing.
- PowerPoint viewers: printing slides and speaker notes on iOS prints the slides or notes, not the viewer page. Printing waits for the print document to load. All five bindings print notes through the shared helper.
- PowerPoint viewers: long-press opens the menus of slide thumbnails, section headers, Slide Sorter tiles, the editing canvas and the slide show on touch screens, where Safari sends no `contextmenu` event. Double-tap renames a section, opens a sorter slide and opens a chart part's format pane.
- PowerPoint viewers: controls that appeared only on hover (version history Restore, media bookmark and assistant chat delete buttons, the embedded object Download and Open bar) are visible on touch screens and on keyboard focus.
- PowerPoint slide shows: the Pixelate effect paints its mosaic in Safari.
- `ooxml-ui` and the PowerPoint viewers: dragging a table border, a crop handle, a ruler or other controls no longer selects page text in Safari. Every `user-select` rule now has its `-webkit-` form.
