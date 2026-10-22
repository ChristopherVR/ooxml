// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word validates property children against their schema sequence; edits must keep that order.
import { isElement, WORD_NS, type XmlElement } from './xml';

/** `CT_RPr` child order (ECMA-376 §17.3.2.28). */
export const RPR_ORDER = [
	'rStyle',
	'rFonts',
	'b',
	'bCs',
	'i',
	'iCs',
	'caps',
	'smallCaps',
	'strike',
	'dstrike',
	'outline',
	'shadow',
	'emboss',
	'imprint',
	'noProof',
	'snapToGrid',
	'vanish',
	'webHidden',
	'color',
	'spacing',
	'w',
	'kern',
	'position',
	'sz',
	'szCs',
	'highlight',
	'u',
	'effect',
	'bdr',
	'shd',
	'fitText',
	'vertAlign',
	'rtl',
	'cs',
	'em',
	'lang',
	'eastAsianLayout',
	'specVanish',
	'oMath',
	'ins',
	'del',
	'moveFrom',
	'moveTo',
	'rPrChange',
];

/**
 * Sorts an element's children into `order`. An entry that lists several names is a choice group
 * whose members may come in any order, so they keep theirs. Elements outside the list
 * (extensions) keep their position relative to the known element before them. Nothing moves when
 * already in order.
 */
export function orderChildren(
	parent: XmlElement,
	order: readonly (string | readonly string[])[],
): void {
	const elements = Array.from(parent.childNodes).filter(isElement);
	const rankOf = (name: string) =>
		order.findIndex((entry) => (typeof entry === 'string' ? entry === name : entry.includes(name)));
	let previousRank = -1;
	const ranked = elements.map((element, index) => {
		const rank = element.namespaceURI === WORD_NS ? rankOf(element.localName ?? '') : -1;
		if (rank >= 0) previousRank = rank;
		return { element, index, rank: rank >= 0 ? rank : previousRank + 0.5 };
	});
	const sorted = [...ranked].sort((a, b) => a.rank - b.rank || a.index - b.index);
	if (sorted.every((entry, index) => entry.element === elements[index])) return;
	for (const { element } of sorted) parent.appendChild(element);
}

/**
 * `CT_SectPr` child order (ECMA-376 §17.6.17). Header and footer references come first, as one
 * choice group (`EG_HdrFtrReferences`): Word interleaves them, the first-page pair last.
 */
export const SECTPR_ORDER = [
	['headerReference', 'footerReference'],
	'footnotePr',
	'endnotePr',
	'type',
	'pgSz',
	'pgMar',
	'paperSrc',
	'pgBorders',
	'lnNumType',
	'pgNumType',
	'cols',
	'formProt',
	'vAlign',
	'noEndnote',
	'titlePg',
	'textDirection',
	'bidi',
	'rtlGutter',
	'docGrid',
	'printerSettings',
	'sectPrChange',
];

/** Sorts `w:sectPr` children into schema order. */
export function orderSectionProperties(sectPr: XmlElement): void {
	orderChildren(sectPr, SECTPR_ORDER);
}
