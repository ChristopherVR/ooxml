# Theming

Excel chrome and native framework controls share the `--office-*` design system with Word,
PowerPoint and Visio. Install the shared theme once, then use a common ancestor to theme
both the canvas and optional controls. Native controls and mounted editors also install it
idempotently when used on their own.

```ts
import { installOfficeUiTheme } from 'ooxml-ui/theme';
installOfficeUiTheme();
```

```html
<section data-office-theme="dark" class="workbook-host">
	<!-- Render the native Toolbar, FormulaBar, SpreadsheetCanvas and SheetTabs here. -->
</section>
```

`data-office-theme` accepts `light`, `dark` and `auto`. It can be set on the document root or
any scoped ancestor. `auto` follows the operating system. A canvas with its default
`theme="auto"` inherits its ancestor's palette. Explicit `theme="light"` or `theme="dark"`
creates a local scope on that editor, keeping the existing property and attribute API.
It does not change sibling host controls or a parent application.

Canonical tokens can be overridden at the same scope:

```css
.workbook-host {
	--office-accent: #1f9d63;
	--office-ring: #1f9d63;
	--office-background: #0f172a;
	--office-surface: #1e293b;
	--office-foreground: #f8fafc;
	--office-border: #64748b;
	--office-radius: 6px;
	--office-font: 'Segoe UI', sans-serif;
}
```

Native optional controls opt into the shared styles with `data-office-block` markers. Their
inputs, buttons, selected states, disabled states and focus rings follow the same tokens.
The stylesheet does not target unmarked custom host controls. Theme changes update CSS in
place, preserving the current engine, active cell, draft, selection and focused input.
Workbook cell colours and the document theme remain unchanged; the sheet paper defaults to
white. `--office-paper` and the legacy `--xve-sheet-bg` are presentation overrides.

## Token overrides

`themeColors` remains supported for per-editor overrides. Each key becomes a `--xve-<kebab-key>` custom property on the host element, with the same token set the Word editor uses (`--dve-*` there), so one design system can drive both:

| Key                                    | Used for                                  |
| -------------------------------------- | ----------------------------------------- |
| `background`, `foreground`             | Editor canvas and text                    |
| `card`, `cardForeground`               | Ribbon and panels                         |
| `popover`, `popoverForeground`         | Menus and dropdowns                       |
| `primary`, `primaryForeground`         | Accent (title bar, selection, active tab) |
| `secondary`, `secondaryForeground`     | Secondary buttons                         |
| `muted`, `mutedForeground`             | Headers, hints, disabled text             |
| `accent`, `accentForeground`           | Hover and pressed states                  |
| `destructive`, `destructiveForeground` | Errors                                    |
| `border`, `input`, `ring`              | Lines, fields and focus rings             |
| `radius`                               | Corner radius                             |

```ts
editor.themeColors = { primary: '#1f9d63', ring: '#1f9d63', radius: '6px' };
```

You can also set the custom properties in CSS:

```css
xlsx-editor {
	--xve-primary: #1f9d63;
}
```

## Following a site theme

The demo follows the VitePress `vitepress-theme-appearance` key in `localStorage` and sets
`data-office-theme` on the document root. Bind that same attribute to your application's
setting so composed controls and editors change together. The editor never reads
`localStorage` itself.

Legacy `--xve-*` CSS overrides remain supported on the editor or an ancestor. They take
precedence over the corresponding canonical token where supplied. Inner shared primitives
receive those resolved values without redeclaring the canonical palette on the editor
host. Existing light/dark preset objects and `themeColors` keys keep their public shape.
