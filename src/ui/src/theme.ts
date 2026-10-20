import { OFFICE_TOKENS, TOUCH } from './tokens';
import { themeBridge } from './theme-bridge';
import {
	BUILDING_BLOCKS_CSS,
	type OfficeUiBlockKind,
	type OfficeUiProduct,
} from './building-block-styles';
export {
	BUILDING_BLOCKS_CSS,
	type OfficeUiBlockKind,
	type OfficeUiProduct,
} from './building-block-styles';

export {
	shadcnBridge,
	themeBridge,
	type ShadcnBridgeOptions,
	type ThemeBridgeMap,
} from './theme-bridge';
export { COMPACT, FOCUS_RING, OFFICE_TOKENS, tok, TOUCH, type OfficeToken } from './tokens';

/**
 * Theme-token foundation. Custom properties cross shadow boundaries, so one `<style>` in the
 * document head themes every control. Every token also has a fallback inside each control, so
 * installing the theme is optional; a host overrides any token on `:root` or on a container.
 */
export const THEME_STYLE_ID = 'office-ui-theme';

const DARK = `
	--office-foreground: #f9fafb;
	--office-muted-foreground: #9ca3af;
	--office-background: #111827;
	--office-surface: #1f2937;
	--office-border: #374151;
	--office-accent: #6366f1;
	--office-ring: #818cf8;
	--office-danger: #f87171;
	--office-selected: #374151;
	--office-notice-background: rgb(120 53 15 / 20%);
	--office-notice-foreground: #fde68a;
	--office-notice-accent: #fbbf24;
	--office-info: #60a5fa;
	--office-warning: #f59e0b;
	--office-teams-brand: #7f85f5;
	--office-teams-brand-hover: #96a0ff;
	--office-teams-brand-subtle: #2f2f4a;
	--office-teams-rail: #141414;
	--office-teams-panel: #1f1f1f;
	--office-teams-bubble: #2d2d2d;
	--office-teams-bubble-own: #2f2f4a;
	--office-teams-divider: #3d3d3d;
	--office-teams-text-subtle: #adadad;
	--office-shadow: 0 6px 20px rgb(0 0 0 / 45%);
	--office-shadow-lg: 0 10px 25px rgb(0 0 0 / 45%);`;

/** Every token's default on `:root`, generated from the token table. */
const LIGHT = Object.entries(OFFICE_TOKENS)
	.map(([name, value]) => `\t${name}: ${value};`)
	.join('\n');

// Mode scopes reset palette values and their dependents, while typography and geometry inherit.
const darkEntries = [...DARK.matchAll(/(--office-[\w-]+): ([^;]+);/g)].map(
	(match) => [match[1]!, match[2]!] as const,
);
const lightPalette = Object.fromEntries(
	darkEntries.map(([name]) => [name, OFFICE_TOKENS[name as keyof typeof OFFICE_TOKENS]]),
);
const darkPalette = Object.fromEntries(darkEntries);
const modeCss = (selector: string, dark: boolean) =>
	themeBridge(selector, {
		...lightPalette,
		...(dark ? darkPalette : {}),
		'--office-accent-foreground': OFFICE_TOKENS['--office-accent-foreground'],
	}) + `${selector} { color-scheme: ${dark ? 'dark' : 'light'}; }`;

export const THEME_CSS = `
:where(:root) {
${LIGHT}
color-scheme: light;
}
${modeCss(':where([data-office-theme="light"], [data-office-theme="auto"])', false)}
${modeCss(':where([data-office-theme="dark"])', true)}
@media (prefers-color-scheme: dark) {
	${modeCss(':where(:root:not([data-office-theme="light"]), [data-office-theme="auto"])', true)}
}
@media ${TOUCH} {
	:where(:root, [data-office-theme]) { --office-target-size: var(--office-target-size-touch); }
}
@media (forced-colors: active) {
	:where(:root, [data-office-theme]) {
		--office-foreground: CanvasText;
		--office-muted-foreground: GrayText;
		--office-background: Canvas;
		--office-surface: Canvas;
		--office-border: CanvasText;
		--office-accent: Highlight;
		--office-accent-foreground: HighlightText;
		--office-ring: Highlight;
		--office-danger: CanvasText;
	}
}
${BUILDING_BLOCKS_CSS}
`;

/** Idempotent. Browser-only; returns false (and does nothing) without a document. */
export function installOfficeUiTheme(doc: Document | undefined = globalThis.document): boolean {
	if (!doc?.head) return false;
	if (!doc.getElementById(THEME_STYLE_ID)) {
		const style = doc.createElement('style');
		style.id = THEME_STYLE_ID;
		style.textContent = THEME_CSS;
		doc.head.append(style);
	}
	return true;
}

/** Marks only the library control's own root. Theme state and custom tokens stay on host scopes. */
export function prepareOfficeUiBlock(
	host: HTMLElement,
	kind: OfficeUiBlockKind,
	product?: OfficeUiProduct,
): void {
	host.setAttribute('data-office-block', kind);
	if (product) host.setAttribute('data-office-product', product);
	installOfficeUiTheme(host.ownerDocument);
}
