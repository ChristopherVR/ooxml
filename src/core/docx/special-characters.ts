// Run content Word stores as elements rather than text: symbols (`w:sym`), positional tabs
// (`w:ptab`) and soft and non-breaking hyphens. The model keeps each as a character in the run's
// text, with the symbol and positional tab on the run, so the writer can restore the elements.
import {
	isStPTabAlignment,
	isStPTabLeader,
	isStPTabRelativeTo,
} from './generated/wml-simple-types';
import type { TextRun } from './model';
import { getW, makeW, named, type XmlDocument, type XmlElement } from './xml';

export const SOFT_HYPHEN = '\u00AD';
export const NON_BREAKING_HYPHEN = '\u2011';

/** The character a symbol shows: Word codes symbol-font characters from U+F000 (`F0FE`). */
export function symbolCharacter(char: string): string {
	const code = /^[0-9a-f]{1,6}$/i.test(char) ? Number.parseInt(char, 16) : Number.NaN;
	if (!Number.isInteger(code) || code > 0x10ffff) return '';
	return String.fromCodePoint(code < 0x100 ? 0xf000 + code : code);
}

/** Whether a run's text is only its symbol, so the whole run is drawn in the symbol's font. */
export function symbolOnly(run: TextRun): boolean {
	const symbol = run.symbol && symbolCharacter(run.symbol.char);
	return Boolean(symbol && run.text && [...run.text].every((character) => character === symbol));
}

/** The one symbol a run's content shows; `null` when it holds different or unreadable symbols. */
export function runSymbol(content: readonly XmlElement[]): TextRun['symbol'] | null | undefined {
	let found: TextRun['symbol'];
	for (const child of content) {
		if (!named(child, 'sym')) continue;
		const symbol = { font: getW(child, 'font') ?? '', char: getW(child, 'char') ?? '' };
		if (!symbolCharacter(symbol.char)) return null;
		if (found && (found.font !== symbol.font || found.char !== symbol.char)) return null;
		found = symbol;
	}
	return found;
}

/** A run's positional tab; `null` when it mixes positional tabs of different kinds or with tabs. */
export function runPositionalTab(
	content: readonly XmlElement[],
): TextRun['positionalTab'] | null | undefined {
	let found: TextRun['positionalTab'];
	let ordinary = false;
	for (const child of content) {
		if (named(child, 'tab')) ordinary = true;
		if (!named(child, 'ptab')) continue;
		const alignment = getW(child, 'alignment');
		const relativeTo = getW(child, 'relativeTo');
		const leader = getW(child, 'leader');
		if (
			!isStPTabAlignment(alignment) ||
			!isStPTabRelativeTo(relativeTo) ||
			(leader !== undefined && !isStPTabLeader(leader))
		)
			return null;
		const tab = { alignment, relativeTo, ...(leader !== undefined && { leader }) };
		if (found && JSON.stringify(found) !== JSON.stringify(tab)) return null;
		found = tab;
	}
	return found && ordinary ? null : found;
}

/** The text a run's special-character element stands for, or undefined for other content. */
export function specialCharacterText(child: XmlElement): string | undefined {
	if (named(child, 'sym')) return symbolCharacter(getW(child, 'char') ?? '');
	if (named(child, 'ptab')) return '\t';
	if (named(child, 'softHyphen')) return SOFT_HYPHEN;
	if (named(child, 'noBreakHyphen')) return NON_BREAKING_HYPHEN;
	return undefined;
}

function element(doc: XmlDocument, local: string, attributes: Record<string, string | undefined>) {
	const node = makeW(doc, local);
	for (const [name, value] of Object.entries(attributes))
		if (value !== undefined) node.setAttributeNS(node.namespaceURI, `w:${name}`, value);
	return node;
}

/**
 * The element for one character of a run's text that Word stores as an element, or undefined for
 * ordinary text: tabs become `w:ptab` when the run's tabs are positional.
 */
export function specialCharacterElement(
	doc: XmlDocument,
	run: TextRun,
	character: string,
): XmlElement | undefined {
	if (character === '\t' && run.positionalTab) {
		const { relativeTo, alignment, leader } = run.positionalTab;
		return element(doc, 'ptab', { relativeTo, alignment, leader });
	}
	if (character === SOFT_HYPHEN) return makeW(doc, 'softHyphen');
	if (character === NON_BREAKING_HYPHEN) return makeW(doc, 'noBreakHyphen');
	if (run.symbol && character === symbolCharacter(run.symbol.char))
		return element(doc, 'sym', { font: run.symbol.font, char: run.symbol.char });
	return undefined;
}
