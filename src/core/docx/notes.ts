// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, Note, Paragraph } from './model';
import { cellSourceBlocks } from './cell-blocks';
import { children, getW, parseXml, type XmlElement } from './xml';

const skippedTypes = new Set(['separator', 'continuationSeparator']);

/** Parses `word/footnotes.xml` or `word/endnotes.xml`, skipping Word's separator marks. */
export function parseNotesPart(
	xml: string,
	kind: 'footnote' | 'endnote',
	parseBlocks: (container: XmlElement, idPrefix: string) => Block[],
): Note[] {
	const document = parseXml(xml);
	const tag = kind === 'footnote' ? 'footnote' : 'endnote';
	const idPrefix = kind === 'footnote' ? 'fn' : 'en';
	const notes: Note[] = [];
	for (const element of children(document.documentElement, tag)) {
		const type = getW(element, 'type');
		if (type && skippedTypes.has(type)) continue;
		const id = getW(element, 'id');
		if (!id) continue;
		notes.push({ id, blocks: parseBlocks(element, `${idPrefix}${id}-`) });
	}
	return notes;
}

/** Separator stories are separate from numbered notes and never receive note numbers. */
export function parseFootnoteSeparators(
	xml: string,
	parseBlocks: (container: XmlElement, idPrefix: string) => Block[],
): { separator?: Note; continuationSeparator?: Note } {
	const root = parseXml(xml).documentElement;
	const result: { separator?: Note; continuationSeparator?: Note } = {};
	for (const element of children(root, 'footnote')) {
		const type = getW(element, 'type');
		if (type !== 'separator' && type !== 'continuationSeparator') continue;
		// Split mixed-content runs so the rule retains its position among surrounding text.
		for (const run of Array.from(element.getElementsByTagNameNS(root.namespaceURI!, 'r'))) {
			const content = Array.from(run.childNodes)
				.filter((child): child is XmlElement => child.nodeType === 1)
				.filter((child) => child.localName !== 'rPr');
			if (content.length < 2 || !content.some((child) => skippedTypes.has(child.localName)))
				continue;
			const props = children(run, 'rPr')[0];
			for (const child of content) {
				const split = run.cloneNode(false) as XmlElement;
				if (props) split.appendChild(props.cloneNode(true));
				split.appendChild(child.cloneNode(true));
				run.parentNode!.insertBefore(split, run);
			}
			run.parentNode!.removeChild(run);
		}
		result[type] = { id: getW(element, 'id') ?? type, blocks: parseBlocks(element, `fn-${type}-`) };
	}
	return result;
}

const romanNumerals: [number, string][] = [
	[1000, 'M'],
	[900, 'CM'],
	[500, 'D'],
	[400, 'CD'],
	[100, 'C'],
	[90, 'XC'],
	[50, 'L'],
	[40, 'XL'],
	[10, 'X'],
	[9, 'IX'],
	[5, 'V'],
	[4, 'IV'],
	[1, 'I'],
];
function toRoman(value: number): string {
	let remaining = Math.max(1, Math.round(value));
	let result = '';
	for (const [amount, numeral] of romanNumerals) {
		while (remaining >= amount) {
			result += numeral;
			remaining -= amount;
		}
	}
	return result || 'I';
}
function toLetters(value: number): string {
	let remaining = Math.max(1, Math.round(value));
	let result = '';
	while (remaining > 0) {
		remaining--;
		result = String.fromCharCode(97 + (remaining % 26)) + result;
		remaining = Math.floor(remaining / 26);
	}
	return result || 'a';
}

/** Renders a 1-based note number in a Word `w:numFmt` display format. */
export function formatNoteNumber(value: number, format: string | undefined): string {
	switch (format) {
		case 'upperRoman':
			return toRoman(value);
		case 'lowerRoman':
			return toRoman(value).toLowerCase();
		case 'upperLetter':
			return toLetters(value).toUpperCase();
		case 'lowerLetter':
			return toLetters(value);
		case 'decimalZero':
			return value < 10 ? `0${value}` : String(value);
		default:
			return String(value);
	}
}

function walkParagraph(
	paragraph: Paragraph,
	kind: 'footnote' | 'endnote',
	order: Map<string, number>,
) {
	for (const run of paragraph.runs) {
		if (run.noteReference?.kind === kind && !order.has(run.noteReference.id))
			order.set(run.noteReference.id, order.size + 1);
	}
}
function walkBlocks(blocks: Block[], kind: 'footnote' | 'endnote', order: Map<string, number>) {
	for (const block of blocks) {
		if (block.type === 'paragraph') walkParagraph(block, kind, order);
		else
			for (const row of block.rows)
				for (const cell of row) walkBlocks(cellSourceBlocks(cell), kind, order);
	}
}

/** Assigns 1-based note numbers by first-reference order, the same order Word displays. */
export function numberNotesInOrder(
	blocks: Block[],
	kind: 'footnote' | 'endnote',
): Map<string, number> {
	const order = new Map<string, number>();
	walkBlocks(blocks, kind, order);
	return order;
}
