# SmartArt elements (`src/ui/src/smartart`)

The SmartArt UI every product shares: the drawing element, the Insert SmartArt picker, the text pane
and the SmartArt Design galleries. The DOM elements draw and report; they never lay a diagram out or
edit one. The model, the node operations, the drawing and the part writers live in `src/core/diagram`
(`diagram/edit` for node and Design commands, `diagram/builtin` for the built-in layouts). Each product
applies an edit through its own undo history: Word as a ProseMirror transaction, Excel as an edit
session step, PowerPoint through its own model.

| Element or export                                          | What it does                                                                                                                                                                                         |
| ---------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `<office-ui-smartart>`                                     | Draws a `DiagramDrawing` (the drawing the producing application cached). Emits `office-smartart-render` with what it approximated (geometry, fills, flattened 3D). It never hides an approximation.  |
| `<office-ui-smartart-picker>`                              | The body of Insert SmartArt: categories, a gallery of layout tiles and the chosen layout's preview. `items-editor` adds a text box of node texts, one per line.                                      |
| `<office-ui-smartart-text-pane>`                           | The outline of node text lines. It turns typing and keys into requests (`office-smartart-*` events) and never edits the diagram itself.                                                              |
| `SmartArtTextPaneController`                               | The wiring every product shares for the text pane: it creates the pane, turns its requests into `diagram/edit` node commands and runs the Design commands (`diagramDesignEdit`) on the current node. |
| `smartArtColorsGalleryState`, `smartArtStylesGalleryState` | Change Colors and SmartArt Styles as `office-ui-gallery` state, from the `diagram` catalogue and the tiles in `tiles.ts`.                                                                            |

## Events

- `<office-ui-smartart-picker>`: `office-smartart-change` on every choice or edit, and
  `office-smartart-activate` when a tile is double-clicked or Enter is pressed on it (the host inserts,
  as OK would). The host sets `categories`, `layouts`, `labels` and optionally `drawPreview`; without a
  preview each tile draws its layout family in `accent`.
- `<office-ui-smartart-text-pane>`: `office-smartart-text-change` (`{ id, text }`) while typing;
  `office-smartart-node-add` (Enter), `-promote` and `-demote` (Shift+Tab and Tab), `-remove`
  (Backspace or Delete in an empty line, unless the line's `canRemove` is false) and `-select` (focus).
  `closable` adds a close button that fires `office-smartart-text-pane-close`. With `tools` it adds the
  header append button (`office-smartart-node-append`), per-line move, add sub-item and style requests
  (`office-smartart-node-move`, `-add-child`, `-style`). The host sets `nodes` again after it applies a
  request, and `selected` to the node to focus.
- `<office-ui-smartart>`: `office-smartart-render` with the approximation report.

The controller's host supplies `nodes()` (the selected diagram's nodes in tree order), `edit(edit)`
(applies one `DiagramNodeEdit` and returns the node it acted on), `setText(id, text)` and
`toggled(open)`. A product that adds nodes calls `select(nodeId)` on the new node.

## Honest limits

Word and Excel insert and edit SmartArt, and desktop Word and Excel rebuild the layout from the data
part, so an edit drops the cached drawing. Preset outlines the element cannot draw are rectangles, gradient and pattern fills
are approximated and 3D is flattened; the render event lists each one.

## Files

- `smartart.ts`, `smartart-svg.ts`: the drawing element and its SVG renderer.
- `smartart-picker.ts`, `smartart-picker.css`: the Insert SmartArt picker.
- `smartart-text-pane.ts`, `smartart-text-pane-tools.ts`, `text-pane-controller.ts`: the text pane and
  its controller.
- `galleries.ts`, `tiles.ts`, `excel-appearance.json`: the Design galleries, their tiles and the Excel
  appearance table.

Each element is a Lit class on `OfficeElement`, as described in `src/ui/README.md` ("How the elements
are built"). The PowerPoint wrappers (`pptx-ui-smartart-text-pane`, `pptx-ui-insert-smartart-dialog`)
are in `src/ui/src/pptx/web-components`.
