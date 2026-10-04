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

Every entry declares an `import` and a `default` export condition, so CommonJS consumers on Node 20.19+ or 22.12+ load it through `require(esm)`.

Importing an entry never touches the DOM (SSR-safe). Elements are defined by the `define*` /
`register*` functions, which are idempotent, no-ops without a `window`, and refuse to coexist
with an incompatible build of this package.

## Features: elements

| Tag                                                   | Attributes / properties                                                                                                                                                                                                                                                                                                   | Events (all bubble, composed)                                                                                                                     |
| ----------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| `office-ui-button`                                    | `label`, `icon`, `command`, `disabled`, `pressed`, `expanded`, `icon-only`, `variant="stacked"`, `keyshortcuts`                                                                                                                                                                                                           | `office-command` `{ command }`                                                                                                                    |
| `office-ui-checkbox`, `office-ui-switch`              | `checked`, `disabled`, `value`, `aria-label`; form-associated                                                                                                                                                                                                                                                             | `input`, `change`                                                                                                                                 |
| `office-ui-select`                                    | `<option>` / `<optgroup>` children (`data-display-label`, `data-description`, font previews) or the `options` property `[{ value, label, disabled? }]`; `value`, `selectedIndex`, `disabled`, `aria-label`, `variant="ribbon-font"                                                                                        | "ribbon-icon"`, `data-font-picker`; `icon`and`custom` slots; arrows, Home, End, PageUp/PageDown, typeahead                                        | `input`, `change` |
| `office-ui-ribbon-group`                              | `label`, `launcher` (dialog launcher command), `launcher-disabled`; default slot                                                                                                                                                                                                                                          | launcher: `office-command` `{ command }`                                                                                                          |
| `office-ui-ribbon-stack`                              | `orientation="horizontal"`; default slot (three-row column of small commands by default)                                                                                                                                                                                                                                  | none                                                                                                                                              |
| `office-ui-menu-button` / `-menu-item`                | button: `label`, `icon`, `command` (split button), `variant`, `disabled`, `keyshortcuts`; item: `command`, `label`, `icon`, `disabled`, `checked`                                                                                                                                                                         | item choice: `office-command` `{ command }`                                                                                                       |
| `office-ui-toolbar`                                   | `aria-label`, `aria-orientation`; arrow-key focus movement                                                                                                                                                                                                                                                                | none                                                                                                                                              |
| `office-ui-dialog`                                    | `open`, `heading`, `dismissible="false"`, `show()`, `close()`; slots default, `footer`                                                                                                                                                                                                                                    | `office-dialog-close` `{ reason }` (cancelable)                                                                                                   |
| `office-ui-status-bar` / `-status-item`               | `label`, `value`, `interactive`, `id`; controlled bar: `state` `{ items?, toggles?, views?, zoom? }` with `collaboration` and `end` slots                                                                                                                                                                                 | `office-status-activate` `{ id }`                                                                                                                 |
| `office-ui-zoom-slider`                               | `value` (percent), `min`, `max`, `step`, `disabled`, `fit` (fit button name), `label`                                                                                                                                                                                                                                     | `input`, `change`; fit: `office-command` `{ command: 'zoom-fit' }`                                                                                |
| `office-ui-tab-strip`                                 | `tabs` `[{ id, label, title? }]`, `selected`, `label`, `previous-label`, `next-label`, `add-label`, `add-disabled`, `add-title`, `disabled`                                                                                                                                                                               | `office-tab-select` `{ id }` (cancelable); add: `office-command` `{ command: 'tab-add' }`                                                         |
| `office-ui-context-menu` / `office-ui-menu-separator` | Children mode: `label`, `openAt(x, y)`, `close()`, `open` over `office-ui-menu-item` and separator children. Controlled mode: `state` `{ x, y, label, items: [{ id, label, separatorBefore?, heading?, danger?, disabled?, checked?, icon? }], markers?, zIndex?, autoFocus? }`; arrows, Home, End and type-ahead in both | children: item `office-command` `{ command }`; controlled: `office-menu-request` `{ id }`, `office-menu-close` `{ reason }` (never closes itself) |
| `office-ui-command-search`                            | `commands` `[{ id, label, keywords?, description?, disabled?, title? }]`, `placeholder`, `label`, `limit` (Office's "Tell me what you want to do")                                                                                                                                                                        | choice: `office-command` `{ command }`                                                                                                            |
| `office-ui-options-dialog`                            | `categories` `[{ id, label, description?, disabled?, sections: [{ id, title, controls }] }]` (controls: `toggle`, `select`, `number`, `text`, each optionally `disabled` with a reason), `values`, `category`, `open`, `heading`, `show()`, `close()`                                                                     | OK: `office-options-change` `{ values, changed }`                                                                                                 |
| `office-ui-account`                                   | `profile` `{ displayName, avatarColor, initial? }`; slots `sign-in` and default (product sections)                                                                                                                                                                                                                        | `office-profile-change` `{ profile }`                                                                                                             |
| `office-ui-ribbon`                                    | Tab row over tab panels. Panels are children with `data-ribbon-tab`, `data-label`, optional `data-tab-keytip`; slots `quick-access`, `search`, `end`; `selected`, `label`, `file-label`, `no-file`, `file-expanded`, `file-keytip`; `focusFile()`, `focusTab()`                                                           | `office-ribbon-select` `{ tab }` (cancelable), `office-ribbon-file`                                                                               |
| `office-ui-backstage`                                 | `items` `[{ id, label, group?: 'footer', disabled?, title? }]`, `open`, `selected`, `label`, `back-label`, `show(page?)`, `close()`; pages are children with `data-backstage-page`                                                                                                                                        | `office-backstage-select` `{ id }` (cancelable), `office-backstage-close` `{ reason }` (cancelable)                                               |
| `office-ui-find-bar`                                  | `open`, `label`, `input-label`, `placeholder`, `maxlength`, `disabled`, `navigation-disabled`, `previous-label`, `next-label`, `close-label`; `value`, `status`, `statusTitle`, `show()`, `close()`                                                                                                                       | `office-find-input` `{ query }`, `office-find-step` `{ direction }`, `office-find-close`                                                          |
| `office-ui-ruler`                                     | `orientation="vertical"`, `origin` (px offset of zero), `scale` (px per unit, default 96), `direction="reverse"`; decorative                                                                                                                                                                                              | none                                                                                                                                              |
| `office-ui-print-preview`                             | `pages` (rendered nodes; an inert clone of the current one is shown), `index`, `label`, `empty-label`                                                                                                                                                                                                                     | `office-print-preview-page` `{ index }`                                                                                                           |
| `office-ui-radio`                                     | `checked`, `disabled`, `value`, `name`; radios of one tag sharing a `name` form a roving-focus group; form-associated                                                                                                                                                                                                     | `input`, `change`                                                                                                                                 |
| `office-ui-search`                                    | `value`, `placeholder`, `disabled`, `aria-label`, `variant="titlebar"`; `focus()`, `select()`                                                                                                                                                                                                                             | `input`, `change`                                                                                                                                 |
| `office-ui-ribbon-toggle`                             | `label`, `checked`, `disabled`, `title`, `command` (a labelled ribbon checkbox row; static `checkboxTag` picks the inner checkbox)                                                                                                                                                                                        | `office-toggle` `{ command, checked }`                                                                                                            |
| `office-ui-dialog-footer`                             | `state` `{ actions: [{ id, label, variant?, icon?, disabled?, busy?, title?, align?, testId? }] }`, `focusAction(id)`                                                                                                                                                                                                     | `office-dialog-footer-request` `{ id }`                                                                                                           |
| `office-ui-toasts`                                    | `state` `{ toasts: [{ id, code?, severity, message }], overflowCount?, labels? }`; hidden when empty, host positions it                                                                                                                                                                                                   | `office-toasts-request` `{ id: 'dismiss', toastId }` or `{ id: 'dismissAll' }`                                                                    |
| `office-ui-read-only-banner`                          | `state` `{ kind, message, passwordPromptOpen?, passwordError?, checkingPassword?, labels? }`                                                                                                                                                                                                                              | `office-read-only-request` (Edit anyway, Dismiss, password submit or cancel)                                                                      |
| `office-ui-paste-options`                             | `state` `{ left, top, options: [{ id, label }], label? }`; fixed 4px past the pasted object                                                                                                                                                                                                                               | `office-paste-options-request` `{ format }`, `office-paste-options-dismiss`                                                                       |
| `office-ui-title-bar`                                 | `state` `{ appMark?, fileName, status?, tone?, autosave?, quickAccess?: { label, items, showLabels? }, search? }`, `placement` (`titleBar`, `belowRibbon`); slots `collaboration`, `account`; tokens `--office-title-bar-*`                                                                                               | `office-autosave-toggle`, `office-command` `{ command }`, `office-command-search` `{ query, command? }`                                           |
| `office-ui-ribbon-section`                            | `groups` `[{ id, label, commands: [{ id, label, icon, title?, size?, badge?, disabled?, active?, compact?, hidden?, caret?, pressed?, expanded?, column? }] }]`; keyed groups and commands in the light DOM                                                                                                               | the commands' own events (`office-command`)                                                                                                       |
| `office-ui-presence`                                  | `participants`, `max`, `label`                                                                                                                                                                                                                                                                                            | `office-presence-select` `{ id }`                                                                                                                 |
| `office-ui-smartart`                                  | `drawing` (core `DiagramDrawing`), `schemeColors`, `label`                                                                                                                                                                                                                                                                | `office-smartart-render` (render report)                                                                                                          |
| `office-ui-icon`                                      | `name`, `label`                                                                                                                                                                                                                                                                                                           | none                                                                                                                                              |

The local profile behind `office-ui-account` is stored on the device only, under one suite-wide
key (`readOfficeProfile`, `writeOfficeProfile`, `clearOfficeProfile`), so every Office viewer on
an origin shows the same name and avatar colour.

`attachKeyTips(scope)` (from `ooxml-ui/controls`) adds Office KeyTips to any element or shadow
root: press and release Alt to show badges on controls marked `data-keytip="H"`, type a badge to
activate it (controls inside open shadow roots of shared elements, such as ribbon tabs, are included), and use `data-keytip-panel` / `data-keytip-level` for a tab's second level. Escape
steps back; disabled controls are dimmed and never run.

Elements with intent events read their event name and test-id prefix from static class fields (`requestEvent`, `testIdPrefix`), so a product can register a subclass under its own tag that keeps its published contract.

Setting a property never emits an event; only user activation does. Controls honour
`prefers-color-scheme`, `forced-colors: active` (system colours and a visible focus ring) and
coarse pointers (44px targets through `--office-target-size`).

## Design tokens

Every control is token based: no colour, length, radius, font size, shadow, duration or layer
is written literally. `OFFICE_TOKENS` (from `ooxml-ui/theme`) is the single table of
`--office-*` tokens and their defaults; controls read them through `tok(name)`, which expands to
`var(--office-…, <default>)`, so they render correctly with or without the installed theme.
`THEME_CSS` declares every default on `:root`, plus dark-mode and forced-colours values. Override
any token on `:root` or on any container to retheme or resize the suite:

- semantic colours: `--office-foreground`, `-background`, `-surface`, `-selected`, `-border`,
  `-accent`, `-accent-foreground`, `-ring`, `-danger`, `-warning`, `-info`, `-popover`, `-overlay`,
  `-notice-*`, `-paper`;
- scales: `--office-space-*`, `--office-radius-*`, `--office-font-size-*`, `--office-icon-size-*`,
  `--office-control-height-*`, `--office-shadow-*`, `--office-duration-*`, `--office-z-*`;
- component tokens that default to the scales, for example `--office-checkbox-size`,
  `--office-switch-width`, `--office-ribbon-tab-height`, `--office-backstage-nav-width` and
  `--office-dialog-max-width`.

Only the two layout breakpoints (`COMPACT`, `TOUCH`) are literal, because CSS does not allow
custom properties in media queries. A test fails when any control's stylesheet contains a raw
colour or length outside a token, or references an undeclared token.

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
