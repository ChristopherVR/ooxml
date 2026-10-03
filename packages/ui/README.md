# ooxml-ui

[![npm version](https://img.shields.io/npm/v/ooxml-ui.svg)](https://www.npmjs.com/package/ooxml-ui)
[![license](https://img.shields.io/npm/l/ooxml-ui.svg)](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/ooxml-ui.svg)](https://www.npmjs.com/package/ooxml-ui)

> Shared Office controls, styles and custom elements for browser applications.

[Try the apps](https://christophervr.github.io/ooxml/) | [npm](https://www.npmjs.com/package/ooxml-ui) | [Full docs](https://christophervr.github.io/ooxml/) | [Source](https://github.com/ChristopherVR/ooxml)

Shared, format-neutral web components and styles for the Office viewers (Word, PowerPoint and
Excel). Vanilla custom elements with shadow-root controls, typed events, no framework and no
runtime dependency except `ooxml-core` (types and a few pure helpers).

> **You do not install this package.** It is a regular `dependency` of every published editor
> package (`docx-*-viewer`, `pptx-*-viewer` and `xlsx-*-viewer`), so installing the editor of your framework pulls it
> in. It is published separately only so that the editors share one copy of the controls. Do not
> add it to your own install instructions or `package.json`; import controls through the editor
> package, which registers them for you. Importing `ooxml-ui` directly is for
> the viewer packages themselves and for people building their own Office-style UI.

## Install

Applications using a viewer receive this package as a dependency. For a custom
Office UI, install it directly:

```bash
npm install ooxml-ui
```

## Quick start

```js
import { registerOfficeUi } from 'ooxml-ui';

registerOfficeUi();
const button = document.createElement('office-ui-button');
button.setAttribute('label', 'Save');
button.setAttribute('command', 'save');
button.addEventListener('office-command', (event) => console.log(event.detail));
document.body.append(button);
```

## API: entries (ESM only)

| Entry               | Contents                                                                                                     |
| ------------------- | ------------------------------------------------------------------------------------------------------------ |
| `ooxml-ui`          | everything below plus `registerOfficeUi()` and `OFFICE_UI_TAGS`                                              |
| `ooxml-ui/theme`    | `THEME_CSS`, `installOfficeUiTheme()` (tokens, dark mode, forced-colors, touch targets)                      |
| `ooxml-ui/icons`    | icon registry (`registerIcon`, `getIcon`, `listIcons`) and `<office-ui-icon>`                                |
| `ooxml-ui/controls` | button, checkbox, switch, select, ribbon group, toolbar, dialog, status bar and item, zoom slider, tab strip |
| `ooxml-ui/presence` | `<office-ui-presence>` avatar stack for collaboration awareness                                              |
| `ooxml-ui/smartart` | `<office-ui-smartart>`: draws a core `DiagramDrawing` as SVG                                                 |

Importing an entry never touches the DOM (SSR-safe). Elements are defined by the `define*` /
`register*` functions, which are idempotent, no-ops without a `window`, and refuse to coexist
with an incompatible build of this package.

## Features: elements

| Tag                                                   | Attributes / properties                                                                                                                           | Events (all bubble, composed)                                                             |
| ----------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| `office-ui-button`                                    | `label`, `icon`, `command`, `disabled`, `pressed`, `expanded`, `icon-only`, `variant="stacked"`, `keyshortcuts`                                   | `office-command` `{ command }`                                                            |
| `office-ui-checkbox`, `office-ui-switch`              | `checked`, `disabled`, `value`, `aria-label`; form-associated                                                                                     | `input`, `change`                                                                         |
| `office-ui-select`                                    | `options` `[{ value, label, disabled? }]`, `value`, `selectedIndex`, `disabled`                                                                   | `input`, `change`                                                                         |
| `office-ui-ribbon-group`                              | `label`, `launcher` (dialog launcher command), `launcher-disabled`; default slot                                                                  | launcher: `office-command` `{ command }`                                                  |
| `office-ui-ribbon-stack`                              | `orientation="horizontal"`; default slot (three-row column of small commands by default)                                                          | none                                                                                      |
| `office-ui-menu-button` / `-menu-item`                | button: `label`, `icon`, `command` (split button), `variant`, `disabled`, `keyshortcuts`; item: `command`, `label`, `icon`, `disabled`, `checked` | item choice: `office-command` `{ command }`                                               |
| `office-ui-toolbar`                                   | `aria-label`, `aria-orientation`; arrow-key focus movement                                                                                        | none                                                                                      |
| `office-ui-dialog`                                    | `open`, `heading`, `dismissible="false"`, `show()`, `close()`; slots default, `footer`                                                            | `office-dialog-close` `{ reason }` (cancelable)                                           |
| `office-ui-status-bar` / `-status-item`               | `label`, `value`, `interactive`, `id`                                                                                                             | `office-status-activate` `{ id }`                                                         |
| `office-ui-zoom-slider`                               | `value` (percent), `min`, `max`, `step`, `disabled`, `fit` (fit button name), `label`                                                             | `input`, `change`; fit: `office-command` `{ command: 'zoom-fit' }`                        |
| `office-ui-tab-strip`                                 | `tabs` `[{ id, label, title? }]`, `selected`, `label`, `previous-label`, `next-label`, `add-label`, `add-disabled`, `add-title`, `disabled`       | `office-tab-select` `{ id }` (cancelable); add: `office-command` `{ command: 'tab-add' }` |
| `office-ui-context-menu` / `office-ui-menu-separator` | `label`; `openAt(x, y)`, `close()`, `open`; `office-ui-menu-item` and separator children                                                          | item choice: `office-command` `{ command }`                                               |
| `office-ui-presence`                                  | `participants`, `max`, `label`                                                                                                                    | `office-presence-select` `{ id }`                                                         |
| `office-ui-smartart`                                  | `drawing` (core `DiagramDrawing`), `schemeColors`, `label`                                                                                        | `office-smartart-render` (render report)                                                  |
| `office-ui-icon`                                      | `name`, `label`                                                                                                                                   | none                                                                                      |

Setting a property never emits an event; only user activation does. Controls honour
`prefers-color-scheme`, `forced-colors: active` (system colours and a visible focus ring) and
coarse pointers (44px targets through `--office-target-size`). Tokens are the `--office-*`
custom properties in `THEME_CSS`; override them on `:root` or any container.

`office-ui-smartart` shows the drawing the producing application cached; it never lays out a
diagram itself. Preset outlines it cannot draw are rendered as rectangles, gradient and pattern
fills are approximated, and 3D is flattened. All of that is listed in the `office-smartart-render`
event detail (`approximatedGeometries`, `approximatedFills`, `flattened3d`); nothing is hidden.

## Scope

Only what every Office product needs lives here. Product-specific ribbon content (PowerPoint
slides, transitions, animations; Word styles and review) stays in the viewers. See
`docs/ooxml-ui-plan.md` in the repository for what moves when, and how the `pptx-ui-*` tags
become aliases of `office-ui-*`.

## Development

```
bun install
bun run --cwd packages/ui typecheck
bun run --cwd packages/ui test
bun run build && bun run --cwd packages/ui build   # core first: the declarations read ../../dist
bun run --cwd packages/ui test:package
```

## Documentation

[Core and UI source](https://github.com/ChristopherVR/ooxml) |
[Shared UI plan](https://github.com/ChristopherVR/ooxml/blob/main/docs/ooxml-ui-plan.md)

## License

Apache-2.0. Several controls are derived from `pptx-viewer`; see
`PROVENANCE.md` at the repository root.
