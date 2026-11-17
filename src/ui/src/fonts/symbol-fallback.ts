import { SYMBOL_FALLBACK_DATA } from './assets/symbol-fallback-data';

export const SYMBOL_FALLBACK_FAMILY = 'OOXML Noto Symbols';
const STYLE_ID = 'ooxml-symbol-fallback';

/** Self-contained font data: also collected by the slide image/print exporter. */
export const SYMBOL_FALLBACK_CSS = `@font-face {
	font-family: "${SYMBOL_FALLBACK_FAMILY}";
	src: url("${SYMBOL_FALLBACK_DATA}") format("woff2");
	font-style: normal;
	font-weight: 400;
	font-display: swap;
	unicode-range: U+2000-2BFF;
}`;

/** Keep the authored face first; CSS selects this face only for missing glyphs. */
export function withSymbolFallback(family: string): string {
	if (typeof document !== 'undefined' && !document.getElementById(STYLE_ID)) {
		const sheet = document.createElement('style');
		sheet.id = STYLE_ID;
		sheet.textContent = SYMBOL_FALLBACK_CSS;
		document.head.appendChild(sheet);
	}
	if (family.includes(`"${SYMBOL_FALLBACK_FAMILY}"`)) return family;
	// Put it before CSS generics: those may already resolve to a system fallback
	// with a tofu box or an unsuitable emoji presentation on iPadOS.
	const generic = /(?:^|,\s*)(?:serif|sans-serif|monospace|cursive|fantasy|system-ui)\s*(?:,|$)/u;
	const at = family.search(generic);
	const fallback = `"${SYMBOL_FALLBACK_FAMILY}"`;
	if (at < 0) return `${family}, ${fallback}`;
	return `${family.slice(0, at)}${at ? ', ' : ''}${fallback}, ${family.slice(at).replace(/^,\s*/u, '')}`;
}
