/**
 * One hyperlink safety policy for every format.
 *
 * - `isSafeHyperlinkHref` is the original strict check (http, https, mailto only) used by
 *   callers that write external relationships from user input; its behaviour is unchanged.
 * - `hyperlinkPolicy` separates two questions: may the href be **stored** in the document
 *   (http, https, mailto, ftp, file, relative paths, UNC and drive paths, in-document
 *   locations such as `#Sheet2!A1`), and may a viewer **open** it on click (http, https and
 *   mailto only). Script-capable schemes (`javascript:`, `vbscript:`, `data:`, `mhtml:` and
 *   relatives) are never stored or opened, including when hidden behind case changes,
 *   whitespace, control characters or zero-width characters.
 *
 * The obfuscation hardening follows `isUrlSafe` of pptx-viewer
 * (`packages/shared/src/render/hyperlink-security.ts`); see PROVENANCE.md.
 */

const SAFE_SCHEMES = new Set(['http:', 'https:', 'mailto:']);

/** Only http(s) and mailto targets are accepted; anything else (notably `javascript:`) is rejected. */
export function isSafeHyperlinkHref(href: string): boolean {
	const scheme = /^([a-z][a-z0-9+.-]*:)/i.exec(href.trim())?.[1];
	return scheme !== undefined && SAFE_SCHEMES.has(scheme.toLowerCase());
}

/** Why a hyperlink may not be stored or opened. */
export type HyperlinkRejection =
	| 'empty'
	| 'script-scheme'
	| 'obfuscated-scheme'
	| 'control-characters'
	| 'unsupported-scheme'
	| 'not-openable';

/** The verdict of {@link hyperlinkPolicy}; `reason` is set whenever `store` or `open` is false. */
export interface HyperlinkPolicy {
	/** The href may be kept in (written to) the document. */
	readonly store: boolean;
	/** A viewer may navigate to the href when the user activates the link. */
	readonly open: boolean;
	readonly reason?: HyperlinkRejection;
}

const OPEN_SCHEMES = new Set(['http', 'https', 'mailto']);
const STORE_SCHEMES = new Set([...OPEN_SCHEMES, 'ftp', 'file']);
// Schemes that run script or render attacker-supplied documents. Listed even though the
// allowlist already rejects them, so the reason reported is the precise one.
const SCRIPT_SCHEMES = new Set([
	'javascript',
	'vbscript',
	'livescript',
	'jscript',
	'mocha',
	'data',
	'mhtml',
	'ms-its',
	'mk',
	'x-javascript',
	'blob',
]);

// Characters a browser's URL parser or a naive check may ignore: ASCII whitespace and
// controls, DEL, soft hyphen, zero-width space/joiners, bidi marks, word joiner and the BOM.
// Built from code points so no invisible character sits in this source file.
const IGNORABLE = new RegExp(
	'[\\u0000-\\u0020\\u007f\\u00ad\\u200b-\\u200f\\u202a-\\u202e\\u2060-\\u2064\\ufeff]',
	'gu',
);
const CONTROL = new RegExp('[\\u0000-\\u001f\\u007f]', 'u');
const SCHEME = /^([a-z][a-z0-9+.-]*):/;

const STORE_ONLY: HyperlinkPolicy = { store: true, open: false, reason: 'not-openable' };
const reject = (reason: HyperlinkRejection): HyperlinkPolicy => ({
	store: false,
	open: false,
	reason,
});

/**
 * Classify a hyperlink target. Leading and trailing whitespace is ignored; anything else
 * that could disguise the scheme makes the href unsafe.
 */
export function hyperlinkPolicy(href: string | undefined | null): HyperlinkPolicy {
	if (typeof href !== 'string') return reject('empty');
	const trimmed = href.trim();
	const normalized = trimmed.replace(IGNORABLE, '').toLowerCase();
	if (normalized.length === 0) return reject('empty');
	const scheme = SCHEME.exec(normalized)?.[1];
	if (scheme !== undefined && SCRIPT_SCHEMES.has(scheme)) return reject('script-scheme');
	if (scheme !== undefined && !trimmed.toLowerCase().startsWith(`${scheme}:`))
		return reject('obfuscated-scheme');
	if (CONTROL.test(trimmed)) return reject('control-characters');
	// A one-letter "scheme" is a Windows drive path (C:\Reports\q1.xlsx).
	if (scheme === undefined || scheme.length === 1) return STORE_ONLY;
	if (OPEN_SCHEMES.has(scheme)) return { store: true, open: true };
	if (STORE_SCHEMES.has(scheme)) return STORE_ONLY;
	return reject('unsupported-scheme');
}

/** True when the href may be written to the document (see {@link hyperlinkPolicy}). */
export function isStorableHyperlinkHref(href: string | undefined | null): boolean {
	return hyperlinkPolicy(href).store;
}

/** True when a viewer may open the href on click (see {@link hyperlinkPolicy}). */
export function isOpenableHyperlinkHref(href: string | undefined | null): boolean {
	return hyperlinkPolicy(href).open;
}
