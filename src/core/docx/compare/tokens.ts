import type { Block, Paragraph, Table, TextRun } from '../model';
import type { CompareGranularity } from './types';

/** A slice of one source run's text; an atomic run is one piece with its whole (empty) text. */
export interface Piece {
	run: TextRun;
	text: string;
}

/**
 * What the comparison aligns. Text tokens are words, single spaces or punctuation (or single
 * characters at character granularity) and may span runs; an atomic token is a run that is not
 * text (a picture, a field marker, a break); a mark is a paragraph mark, owned by the paragraph
 * it ends; a table is compared as a whole block.
 */
export type Token =
	| { kind: 'text'; key: string; pieces: Piece[] }
	| { kind: 'atomic'; key: string; run: TextRun }
	| { kind: 'mark'; key: string; paragraph: Paragraph }
	| { kind: 'table'; key: string; table: Table };

/** The key of a paragraph mark; the story's last mark gets {@link END_KEY}. */
export const MARK_KEY = ' ';
/** The last paragraph mark of a story, which Word never marks as inserted or deleted. */
export const END_KEY = '\u0003end';

const WORD = /[\p{L}\p{N}\p{M}]+(?:['’][\p{L}\p{N}\p{M}]+)*|\s|./gsu;

/** Runs that are not text: compared as a whole, never split. */
export function isAtomicRun(run: TextRun): boolean {
	return Boolean(
		run.image ||
		run.equation ||
		run.fieldChar ||
		run.fieldCode !== undefined ||
		run.noteReference ||
		run.noteMark ||
		run.break ||
		run.noteSeparator,
	);
}

// Only identity-bearing fields: ids and relationship names differ between two saves of one picture.
const ATOMIC_FIELDS = [
	'image',
	'partName',
	'widthPx',
	'heightPx',
	'altText',
	'title',
	'textBoxText',
	'equation',
	'omml',
	'display',
	'fieldChar',
	'fieldCode',
	'noteReference',
	'kind',
	'noteMark',
	'break',
	'noteSeparator',
];
const atomicKey = (run: TextRun): string => `￼${JSON.stringify(run, ATOMIC_FIELDS)}`;

function splitText(text: string, granularity: CompareGranularity): string[] {
	return granularity === 'character' ? Array.from(text) : (text.match(WORD) ?? []);
}

/** Text tokens of consecutive text runs: tokens may cross run boundaries. */
function textTokens(runs: TextRun[], granularity: CompareGranularity): Token[] {
	const text = runs.map((run) => run.text).join('');
	const tokens: Token[] = [];
	let runIndex = 0;
	let offset = 0;
	for (const part of splitText(text, granularity)) {
		const pieces: Piece[] = [];
		let remaining = part.length;
		while (remaining > 0) {
			const run = runs[runIndex]!;
			const take = Math.min(remaining, run.text.length - offset);
			if (take > 0) pieces.push({ run, text: run.text.slice(offset, offset + take) });
			remaining -= take;
			offset += take;
			if (offset >= run.text.length) {
				runIndex++;
				offset = 0;
			}
		}
		tokens.push({ kind: 'text', key: part, pieces });
	}
	return tokens;
}

/** A paragraph's content tokens, without its mark. */
export function paragraphTokens(paragraph: Paragraph, granularity: CompareGranularity): Token[] {
	const tokens: Token[] = [];
	let pending: TextRun[] = [];
	const flush = () => {
		if (pending.length) tokens.push(...textTokens(pending, granularity));
		pending = [];
	};
	for (const run of paragraph.runs) {
		// An empty run (a tab-less marker the model keeps, say) is kept as a token of its own, so a
		// paragraph written back keeps every run it had.
		if (isAtomicRun(run) || !run.text) {
			flush();
			tokens.push({ kind: 'atomic', key: atomicKey(run), run });
		} else pending.push(run);
	}
	flush();
	return tokens;
}

/** The text of a paragraph as the comparison sees it, for aligning whole paragraphs. */
export function paragraphKey(paragraph: Paragraph, granularity: CompareGranularity): string {
	return paragraphTokens(paragraph, granularity)
		.map((token) => token.key)
		.join('');
}

/** A table's text, row by row and cell by cell, for aligning whole tables and rows. */
export function rowKey(row: Table['rows'][number], granularity: CompareGranularity): string {
	return row
		.map((cell) =>
			cell.paragraphs.map((paragraph) => paragraphKey(paragraph, granularity)).join(MARK_KEY),
		)
		.join('\u0007');
}
export function tableKey(table: Table, granularity: CompareGranularity): string {
	return `\u0001table:${table.rows.map((row) => rowKey(row, granularity)).join('\u0006')}`;
}

/** The key aligning a block as a whole: its text, or a table's text. */
export function blockKey(block: Block, granularity: CompareGranularity): string {
	return block.type === 'table'
		? tableKey(block, granularity)
		: `\u0002${paragraphKey(block, granularity)}`;
}

/** The mark token of a paragraph; `last` is the story's final paragraph. */
export function markToken(paragraph: Paragraph, last: boolean): Token {
	return { kind: 'mark', key: last ? END_KEY : MARK_KEY, paragraph };
}

/** Whether every token is whitespace: such a short equal stretch between two changes joins them. */
export const isSpaceToken = (token: Token): boolean =>
	token.kind === 'text' && /^\s+$/u.test(token.key);
