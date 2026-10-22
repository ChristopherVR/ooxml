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

/**
 * The character a symbol shows: Word codes symbol-font characters from U+F000 (`F0FE`). Empty
 * for a code that is not a character (beyond U+10FFFF, or half of a surrogate pair).
 */
export function symbolCharacter(char: string): string {
	const code = /^[0-9a-f]{1,6}$/i.test(char) ? Number.parseInt(char, 16) : Number.NaN;
	if (!Number.isInteger(code) || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return '';
	return String.fromCodePoint(code < 0x100 ? 0xf000 + code : code);
}

/** Whether a run's text is only its symbol, so the whole run is drawn in the symbol's font. */
export function symbolOnly(run: TextRun): boolean {
	const symbol = run.symbol && symbolCharacter(run.symbol.char);
	return Boolean(symbol && run.text && [...run.text].every((character) => character === symbol));
}

/**
 * The run split into stretches of its symbol characters and of other text, so each stretch of
 * symbols can be drawn in the symbol's font and the text in the run's own.
 */
export function symbolSegments(run: TextRun): TextRun[] {
	const symbol = run.symbol && symbolCharacter(run.symbol.char);
	if (!symbol || symbolOnly(run) || !run.text.includes(symbol)) return [run];
	const segments: TextRun[] = [];
	let piece = '';
	let symbols = false;
	for (const character of run.text) {
		const isSymbol = character === symbol;
		if (piece && isSymbol !== symbols) {
			segments.push({ ...run, text: piece });
			piece = '';
		}
		piece += character;
		symbols = isSymbol;
	}
	if (piece) segments.push({ ...run, text: piece });
	return segments;
}

/**
 * Whether a symbol's character is one a symbol font's code maps to (U+F000 to U+F0FF), which is
 * what Word writes for a character inserted from a symbol font, and ordinary text does not hold.
 */
const isSymbolFontCharacter = (character: string): boolean => /^[\uF000-\uF0FF]$/.test(character);

/**
 * The one symbol a run's content shows; `null` when it holds different or unreadable symbols, or
 * one the run's text could not be told apart from. The model keeps a symbol as its character in
 * the run's text, written back as `w:sym` wherever it occurs, so that character must be a symbol
 * font's code that the run's own text does not use (an en dash coded `2013` would also turn the
 * run's ordinary en dashes into symbols).
 */
export function runSymbol(content: readonly XmlElement[]): TextRun['symbol'] | null | undefined {
	let found: TextRun['symbol'];
	for (const child of content) {
		if (!named(child, 'sym')) continue;
		const symbol = { font: getW(child, 'font') ?? '', char: getW(child, 'char') ?? '' };
		if (!isSymbolFontCharacter(symbolCharacter(symbol.char))) return null;
		if (found && (found.font !== symbol.font || found.char !== symbol.char)) return null;
		found = symbol;
	}
	const character = found && symbolCharacter(found.char);
	if (
		character &&
		content.some(
			(child) =>
				(named(child, 't') || named(child, 'delText')) && child.textContent?.includes(character),
		)
	)
		return null;
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
