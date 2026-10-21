import { legacyThemeBridge, LEGACY_OFFICE_THEME_TOKENS } from '../../theme';
import { withOfficeThemeCss } from '../theme/office-css';
import { STATUS_BAR_METRICS, TITLE_BAR_METRICS as M } from '../render';
import { tok } from './control-tokens';

/**
 * The `pptx-ui-*` tags implemented by `ooxml-ui` (see `office-aliases.ts`). Their shadow CSS reads
 * only `--office-*` tokens; this bridge feeds those from the viewer's `--pptx-*` theme and control
 * tokens. It is declared on the hosts themselves because pptx colours resolve where they are read
 * (the `--pptx-*` theme lives on the viewer root, not `:root`), so a viewer theme, a title-bar
 * override of `--pptx-switch-*` or any host override keeps working unchanged.
 */
export const OFFICE_ALIAS_TAGS = [
	'pptx-ui-checkbox',
	'pptx-ui-switch',
	'pptx-ui-radio',
	'pptx-ui-search',
	'pptx-ui-color-grid',
	'pptx-ui-select',
	'pptx-ui-ribbon-toggle',
	'pptx-ui-dialog-footer',
	'pptx-ui-compat-toasts',
	'pptx-ui-read-only-banner',
	'pptx-ui-paste-options',
	'pptx-ui-ribbon-command',
	'pptx-ui-ribbon-group',
	'pptx-ui-context-menu',
	'pptx-ui-title-bar',
	'pptx-ui-status-bar',
	'pptx-ui-ribbon-gallery',
	'pptx-ui-ribbon-actions',
	'pptx-ui-options-dialog',
] as const;

/** Every alias maps the shared tokens onto the pptx theme and control tokens. */
const BASE = `
	--office-ring: var(--pptx-focus-ring-color, var(--_pptx-ring));
	--office-focus-width: ${tok('--pptx-focus-ring-width')};
	--office-focus-offset: ${tok('--pptx-focus-ring-offset')};
	--office-target-size-touch: ${tok('--pptx-touch-target')};
	--office-field-height: ${tok('--pptx-field-height')};
	--office-field-height-lg: ${tok('--pptx-field-height-lg')};
	--office-field-radius: ${tok('--pptx-field-radius')};
	--office-field-border: ${tok('--pptx-field-border')};
	--office-field-border-focus: ${tok('--pptx-field-border-focus')};
	--office-field-background: ${tok('--pptx-field-bg')};
	--office-field-foreground: ${tok('--pptx-field-fg')};
	--office-field-placeholder: ${tok('--pptx-field-placeholder')};
	--office-checkbox-size: ${tok('--pptx-checkbox-size')};
	--office-checkbox-size-touch: ${tok('--pptx-checkbox-size-touch')};
	--office-checkbox-radius: ${tok('--pptx-checkbox-radius')};
	--office-checkbox-border: ${tok('--pptx-checkbox-border')};
	--office-checkbox-background: ${tok('--pptx-checkbox-bg')};
	--office-checkbox-accent: ${tok('--pptx-checkbox-accent')};
	--office-checkbox-accent-foreground: ${tok('--pptx-checkbox-accent-fg')};
	--office-radio-dot-size: ${tok('--pptx-radio-dot-size')};
	--office-radio-dot-size-touch: ${tok('--pptx-radio-dot-size-touch')};
	--office-switch-width: ${tok('--pptx-switch-width')};
	--office-switch-height: ${tok('--pptx-switch-height')};
	--office-switch-width-touch: ${tok('--pptx-switch-width')};
	--office-switch-height-touch: ${tok('--pptx-switch-height')};
	--office-switch-border-width: 0px;
	--office-switch-knob-size: ${tok('--pptx-switch-knob-size')};
	--office-switch-knob-size-touch: ${tok('--pptx-switch-knob-size')};
	--office-switch-knob-offset: ${tok('--pptx-switch-knob-offset')};
	--office-switch-knob-travel: ${tok('--pptx-switch-knob-travel')};
	--office-switch-track: ${tok('--pptx-switch-track')};
	--office-switch-track-on: ${tok('--pptx-switch-track-on')};
	--office-switch-thumb: ${tok('--pptx-switch-thumb')};
	--office-switch-thumb-on: ${tok('--pptx-switch-thumb')};
`;

/**
 * The pptx look of the shared title bar, from the shared chrome metrics: `--office-title-bar-*`
 * tokens, the compact AutoSave switch (its `--pptx-switch-*` inputs inherit into the nested
 * `pptx-ui-switch`) and hiding both placements on phones, where the compact toolbar covers them.
 */
const TITLE_BAR = `:host {
	--office-title-bar-height: ${M.height}px;
	--office-title-bar-background: color-mix(in srgb, var(--pptx-secondary, #1e1e2e) 80%, transparent);
	--office-title-bar-foreground: var(--pptx-card-foreground, #e2e8f0);
	--office-title-bar-name-foreground: var(--pptx-foreground, #f1f5f9);
	--office-title-bar-border: color-mix(in srgb, var(--pptx-border, #33334d) 60%, transparent);
	--office-title-bar-mark-size: ${M.logoSize}px;
	--office-title-bar-mark-background: ${M.logoBackground};
	--office-title-bar-mark-foreground: #fff;
	--office-title-bar-result-selected: var(--pptx-accent, #33334d);
	--office-font-size-xs: ${M.fontSize}px;
	--office-font-size-sm: ${M.fileNameFontSize}px;
	--office-font-size-2xs: ${M.logoFontSize}px;
	--office-font-weight-medium: ${M.fileNameFontWeight};
	--office-icon-size: 14px;
	--office-icon-stroke: 1.4;
	--office-control-height-md: 24px;
	--office-selected: color-mix(in srgb, var(--pptx-accent, #33334d) 60%, transparent);
	--office-popover: var(--pptx-popover, var(--pptx-card, #1e1e2e));
	--pptx-switch-width: ${M.switchTrackWidth}px;
	--pptx-switch-height: ${M.switchTrackHeight}px;
	--pptx-switch-knob-size: ${M.switchKnobSize}px;
	--pptx-switch-knob-offset: ${M.switchKnobOffsetOff}px;
	--pptx-switch-knob-travel: ${M.switchKnobOffsetOn - M.switchKnobOffsetOff}px;
	--pptx-switch-track: color-mix(in srgb, var(--pptx-muted-foreground, #a5a5b5) 40%, transparent);
	--pptx-switch-thumb: #fff;
}
@media (max-width: 767px), (max-width: 1023px) and (max-height: 520px) { :host { display: none !important; } }`;

/** The pptx look of the shared status bar: tokens only, sized by the shared chrome metrics. */
const STATUS_BAR = `:host {
	--office-status-bar-height: ${STATUS_BAR_METRICS.height}px;
	--office-status-bar-separator-opacity: 0.6;
	--office-surface: color-mix(in srgb, var(--pptx-secondary, #1e1e2e) 50%, transparent);
	--office-border: var(--pptx-border, #33334d);
	--office-foreground: var(--pptx-card-foreground, #e2e8f0);
	--office-font-size-xs: 10px;
	--office-control-height-xs: 24px;
	--office-icon-stroke: 1.45;
	line-height: 1.2;
}`;

/**
 * Per-tag extras: pptx sizes and colours where its controls differ from the Office defaults.
 * Only `:host` rules, so no product style depends on a shared element's internal class names.
 */
const EXTRAS: Partial<Record<(typeof OFFICE_ALIAS_TAGS)[number], string>> = {
	'pptx-ui-title-bar': TITLE_BAR,
	'pptx-ui-status-bar': STATUS_BAR,
	'pptx-ui-ribbon-gallery': `:host { --office-foreground: var(--pptx-foreground, #f9fafb); --office-background: var(--pptx-background, #111827);
	--office-gallery-strip-background: color-mix(in srgb, var(--pptx-muted, #2a2a3d) 30%, transparent);
	--office-command-icon-color: var(--pptx-primary, #6366f1); --office-z-popover: 1200;
	--office-popover-foreground: var(--pptx-popover-foreground, #f9fafb);
	--office-icon-stroke: 1.6; --office-line-height-tight: 11px; }`,
	'pptx-ui-select': `:host { --office-select-background: var(--pptx-select-control-bg, var(--pptx-background, #11111b)); }
:host(.bg-muted) { --office-select-background: var(--pptx-muted, #1f2937); }
:host(.bg-popover) { --office-select-background: var(--pptx-popover, #111827); }`,
	'pptx-ui-ribbon-command': `:host { --office-command-icon-color: var(--pptx-primary, #6366f1); --office-font-size-sm: 12px;
	--office-line-height-tight: 1.25; }`,
	'pptx-ui-ribbon-group': `:host { --office-command-icon-color: var(--pptx-primary, #6366f1);
	--office-ribbon-group-height: 84px; --office-ribbon-group-padding: 3px 6px 0; --office-ribbon-separator-inset: 6px;
	--office-ribbon-separator: color-mix(in srgb, var(--pptx-border, #33334d) 80%, transparent);
	--office-ribbon-caption-size: 11px; --office-ribbon-caption-line-height: 16px; --office-ribbon-caption-padding: 0 14px;
	--office-launcher-inset-end: -2px; --office-launcher-inset-bottom: 1px; --office-launcher-width: 14px;
	--office-launcher-height: 14px; --office-ribbon-collapse-x: var(--pptx-collapse-x, 8px);
	--office-ribbon-collapse-y: var(--pptx-collapse-y, 120px); }`,
	// The tab row's Comments and Share at the pptx tab-row scale (11px labels, 24px controls).
	'pptx-ui-ribbon-actions': `:host { --office-font-size: 11px; --office-control-height-md: 24px;
	--office-icon-size-md: 14px; --office-font-size-2xs: 8px; padding-inline-start: 0; }`,
	// File > Options keeps its wide footprint (56rem, a 176px rail) rather than Office's 728px.
	'pptx-ui-options-dialog': `:host { --office-options-width: min(56rem, calc(100vw - 32px));
	--office-options-nav-width: 176px; --office-options-height: 540px; --office-z-dialog: 1200; }`,
	// The inspector's Theme Colors keep the pptx 2px between swatches, so they stay legible in a
	// half-width field.
	'pptx-ui-color-grid': `:host { --office-color-swatch-gap: 2px; }`,
	'pptx-ui-context-menu': `:host { --office-menu-min-width: 180px; --office-danger: var(--pptx-destructive, #f87171);
	--office-selected: var(--pptx-accent, #33334d); --office-radius-md: var(--pptx-radius, 6px); }`,
};

/**
 * The bridge for one alias, as `:host` rules in its own shadow root. It must travel with the
 * element: controls rendered inside another shadow root (the title bar's search and AutoSave
 * switch) never see document styles.
 */
export function bridgeCss(tag: (typeof OFFICE_ALIAS_TAGS)[number]): string {
	const extras = EXTRAS[tag] ?? '';
	// Basic canonical colors belong on inner chrome. Their inherited host inputs remain readable.
	const palette = new Set<string>(Object.values(LEGACY_OFFICE_THEME_TOKENS));
	const inner = extras.replace(
		/(:host(?:\([^)]*\))?)\s*\{([^{}]*)\}/g,
		(_rule, selector: string, body: string) => {
			const host: string[] = [];
			const children: string[] = [];
			for (const declaration of body
				.split(';')
				.map((value) => value.trim())
				.filter(Boolean)) {
				const property = declaration.slice(0, declaration.indexOf(':')).trim();
				if (palette.has(property)) children.push(declaration);
				else {
					host.push(declaration);
					if (property.startsWith('--office-')) {
						// Resolve product roles on the host, where inline overrides win, before the inner
						// theme bridge declares its derived defaults (for example the app mark color).
						const input = `--_pptx-control-${property.slice('--office-'.length)}`;
						host.push(`${input}: var(${property})`);
						children.push(`${property}: var(${input})`);
					}
				}
			}
			return `${selector} { ${host.join(';')}; } ${selector} > * { ${children.join(';')}; }`;
		},
	);
	const hostTokens = BASE.replace(/--office-ring:[^;]+;/, '');
	const hostColor = ['pptx-ui-checkbox', 'pptx-ui-radio', 'pptx-ui-switch'].includes(tag)
		? ''
		: 'color:var(--_pptx-foreground);';
	return withOfficeThemeCss(
		`:host { ${hostTokens} ${hostColor} font-family:var(--_pptx-font); }\n${legacyThemeBridge(':host > *', '--_pptx-')}\n:host > * {${BASE}}\n${inner}\n:host(:focus-visible) { outline-color:var(--pptx-focus-ring-color,var(--_pptx-ring)); }`,
		':host',
	);
}

/**
 * Document-level rules for alias hosts. The token mapping itself lives in each element
 * (`bridgeCss`); this keeps utility classes a binding puts on a select host from drawing a box.
 */
export const OFFICE_TOKEN_BRIDGE = `
/* Bindings put utility classes on the host; only the shadow trigger draws a box. */
pptx-ui-select { border: 0 !important; padding: 0 !important; background: transparent !important; box-shadow: none !important; }
`;
