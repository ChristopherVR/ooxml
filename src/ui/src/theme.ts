import { CARDS_LOOK_CSS } from './look';
import { LOOK_DARK_TOKENS } from './look-tokens';
import { isDerivedToken, OFFICE_TOKENS, TOUCH, type OfficeToken } from './tokens';
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
	LEGACY_OFFICE_THEME_TOKENS,
	legacyThemeAliases,
	legacyThemeBridge,
	shadcnBridge,
	themeBridge,
	type ShadcnBridgeOptions,
	type ThemeBridgeMap,
} from './theme-bridge';
export {
	applyOfficeLook,
	CARDS_LOOK_CSS,
	CARDS_LOOK_PROPERTIES,
	normalizeOfficeLook,
	OFFICE_LOOK_ATTRIBUTE,
	OFFICE_LOOKS,
	type OfficeLook,
} from './look';
export { LOOK_DARK_TOKENS, LOOK_TOKENS } from './look-tokens';
export {
	COMPACT,
	FOCUS_RING,
	isDerivedToken,
	OFFICE_TOKENS,
	tok,
	TOUCH,
	type OfficeToken,
} from './tokens';

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
	--office-color-scheme: dark;
	--office-ring: #818cf8;
	--office-danger: #f87171;
	--office-selected: #374151;
	--office-notice-background: rgb(120 53 15 / 20%);
	--office-notice-foreground: #fde68a;
	--office-notice-accent: #fbbf24;
	--office-info: #60a5fa;
	--office-warning: #f59e0b;
	--office-teams-brand: #7f85f5;
	--office-teams-brand-subtle: #2f2f4a;
	--office-teams-rail: #141414;
	--office-teams-panel: #1f1f1f;
	--office-teams-bubble: #2d2d2d;
	--office-teams-bubble-own: #2f2f4a;
	--office-teams-divider: #3d3d3d;
	--office-teams-text-subtle: #adadad;
	--office-shadow: 0 6px 20px rgb(0 0 0 / 45%);
	--office-shadow-lg: 0 10px 25px rgb(0 0 0 / 45%);
${Object.entries(LOOK_DARK_TOKENS)
	.map(([name, value]) => `\t${name}: ${value};`)
	.join('\n')}`;

/**
 * Every literal default on `:root`, generated from the token table. Derived tokens stay out: a
 * custom property resolves its `var()` where it is declared, so declaring one here would freeze it
 * to the page's colours and a container's override of the token it reads would never reach it.
 */
const LIGHT = Object.entries(OFFICE_TOKENS)
	.filter(([name]) => !isDerivedToken(name as OfficeToken))
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

/**
 * Locally installed CJK UI faces per language, after the Latin ones. Han characters are shared by
 * Chinese, Japanese and Korean but drawn differently, so the faces are keyed on the host's `lang`
 * (a document or container attribute) instead of being appended to every stack. No web fonts.
 */
export const CJK_UI_FONTS: Readonly<Record<'ja' | 'ko' | 'zh-CN', string>> = {
	ja: '"Yu Gothic UI", "Meiryo UI", Meiryo, "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Noto Sans CJK JP", "Noto Sans JP"',
	ko: '"Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans CJK KR", "Noto Sans KR"',
	'zh-CN':
		'"Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "Noto Sans SC"',
};
const CJK_LANG_SELECTORS: Readonly<Record<keyof typeof CJK_UI_FONTS, string>> = {
	ja: ':root:lang(ja), [lang|="ja"]',
	ko: ':root:lang(ko), [lang|="ko"]',
	'zh-CN': ':root:lang(zh-CN), :root:lang(zh-Hans), [lang|="zh-CN"], [lang|="zh-Hans"]',
};
const LATIN_UI_FONTS = OFFICE_TOKENS['--office-font'].replace(/,\s*sans-serif$/, '');
const CJK_FONT_CSS = Object.entries(CJK_UI_FONTS)
	.map(
		([lang, faces]) =>
			`:where(${CJK_LANG_SELECTORS[lang as keyof typeof CJK_UI_FONTS]}) { --office-font: ${LATIN_UI_FONTS}, ${faces}, sans-serif; }`,
	)
	.join('\n');

export const THEME_CSS = `
:where(:root) {
${LIGHT}
color-scheme: light;
}
${CJK_FONT_CSS}
${modeCss(':where([data-office-theme="light"], [data-office-theme="auto"])', false)}
${modeCss(':where([data-office-theme="dark"])', true)}
@media (prefers-color-scheme: dark) {
	${modeCss(':where(:root:not([data-office-theme="light"]), [data-office-theme="auto"])', true)}
}
@media ${TOUCH} {
	:where(:root) { --office-target-size: var(--office-target-size-touch); }
}
@media (forced-colors: active) {
	:where(:root, [data-office-theme]) {
		--office-foreground: CanvasText;
		--office-muted-foreground: GrayText;
		--office-background: Canvas;
		--office-surface: Canvas;
		--office-border: CanvasText;
		--office-control-border: CanvasText;
		--office-accent: Highlight;
		--office-accent-foreground: HighlightText;
		--office-ring: Highlight;
		--office-danger: CanvasText;
		--office-card-border: CanvasText;
	}
}
${BUILDING_BLOCKS_CSS}
${CARDS_LOOK_CSS}`;

/** Idempotent per document or shadow root. Browser-only; false without a style container. */
export function installOfficeUiTheme(
	root: Document | ShadowRoot | undefined = globalThis.document,
): boolean {
	if (!root) return false;
	const documentRoot = root.nodeType === 9;
	if (!documentRoot && !('host' in root)) return false;
	const doc = documentRoot ? (root as Document) : root.ownerDocument;
	const container = documentRoot ? (root as Document).head : root;
	if (!doc || !container) return false;
	if (!root.getElementById(THEME_STYLE_ID)) {
		const style = doc.createElement('style');
		style.id = THEME_STYLE_ID;
		style.textContent = THEME_CSS;
		container.append(style);
	}
	return true;
}

function installOfficeUiThemeRoots(host: HTMLElement): boolean {
	let installed = installOfficeUiTheme(host.ownerDocument);
	let root = host.getRootNode();
	while (root.nodeType === 11 && 'host' in root) {
		const shadow = root as ShadowRoot;
		installed = installOfficeUiTheme(shadow) && installed;
		root = shadow.host.getRootNode();
	}
	return installed;
}

/** Install enclosing root rules now, then recheck after a native framework inserts a prepared host. */
export function installOfficeUiThemeFor(host: HTMLElement): boolean {
	const installed = installOfficeUiThemeRoots(host);
	queueMicrotask(() => {
		installOfficeUiThemeRoots(host);
	});
	return installed;
}

/** Marks only the library control's own root. Theme state and custom tokens stay on host scopes. */
export function prepareOfficeUiBlock(
	host: HTMLElement,
	kind: OfficeUiBlockKind,
	product?: OfficeUiProduct,
): void {
	host.setAttribute('data-office-block', kind);
	if (product) host.setAttribute('data-office-product', product);
	installOfficeUiThemeFor(host);
}
