const SAFE_SCHEMES = new Set(['http:', 'https:', 'mailto:']);

/** Only http(s) and mailto targets are accepted; anything else (notably `javascript:`) is rejected. */
export function isSafeHyperlinkHref(href: string): boolean {
	const scheme = /^([a-z][a-z0-9+.-]*:)/i.exec(href.trim())?.[1];
	return scheme !== undefined && SAFE_SCHEMES.has(scheme.toLowerCase());
}
