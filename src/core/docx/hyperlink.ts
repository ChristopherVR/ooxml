// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Resolves w:hyperlink targets (relationship-based external links, or w:anchor internal links)
// and simple, non-field-code HYPERLINK fields (w:fldSimple).
import type { HyperlinkInfo } from './model';
import type { Relationship } from './package-parts';
import { getR, getW, type XmlElement } from './xml';
import { hyperlinkPolicy } from '../opc/safe-href';

/** Schemes Word follows on Ctrl+click beyond the shared http, https and mailto set. */
export const WORD_EXTRA_LINK_SCHEMES: readonly string[] = ['tel', 'ftp'];

/**
 * True for an external link target the editor may write and follow on a user gesture (http,
 * https, mailto, tel and ftp), after the shared policy rejects script and disguised schemes.
 */
export function isWordNavigableHref(href: unknown): href is string {
	return (
		typeof href === 'string' &&
		hyperlinkPolicy(href, { additionalOpenSchemes: WORD_EXTRA_LINK_SCHEMES }).open
	);
}

/** Resolves a `w:hyperlink` element's target using the document's relationship map. */
export function resolveHyperlink(
	element: XmlElement,
	rels: ReadonlyMap<string, Relationship>,
): HyperlinkInfo {
	const info: HyperlinkInfo = {};
	const relId = getR(element, 'id');
	if (relId) {
		const rel = rels.get(relId);
		if (rel && rel.mode === 'External') info.href = rel.target;
	}
	const anchor = getW(element, 'anchor');
	if (anchor) info.anchor = anchor;
	const tooltip = getW(element, 'tooltip');
	if (tooltip) info.tooltip = tooltip;
	return info;
}

/** Parses a literal `HYPERLINK "target"` (optionally `\l anchor`) instruction from `w:fldSimple/@w:instr`. */
export function parseSimpleHyperlinkField(instr: string): HyperlinkInfo | undefined {
	const match = /^\s*HYPERLINK\s+"([^"]*)"(?:\s+\\l\s+"([^"]*)")?/i.exec(instr);
	if (!match) return undefined;
	const [, target, anchor] = match;
	if (anchor) return { anchor };
	if (target) return { href: target };
	return undefined;
}
