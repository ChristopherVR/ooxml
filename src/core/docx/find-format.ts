// Formatted Find over DOCX blocks and document stories: supports text + format search (every
// character of the match must satisfy the criteria) and format-only search (empty query + format
// criteria, matching maximal contiguous formatted spans within each paragraph, as Word does).
import { findMatches, type FindMatch, type FindOptions } from './find-pattern';
import {
	hasFormatCriteria,
	matchesFormatCriteria,
	matchesParagraphFormat,
	type FindFormatContext,
	type FindFormatCriteria,
} from './find-format-criteria';
import { findStories, findStoryKey } from './find-stories';
import { FIND_LINE_BREAK, FIND_OBJECT, FIND_PAGE_BREAK, FIND_PARAGRAPH_MARK } from './find-syntax';
import type { Block, DocumentModel, Paragraph, TextRun } from './model';

/** Options for formatted search over DOCX blocks or a `DocumentModel`. */
export interface FindFormatOptions extends FindOptions {
	/** Style and theme catalogs for resolving inherited run/paragraph formatting. */
	readonly context?: FindFormatContext;
}

/** A Find match carrying both flat-text offsets and exact paragraph/run coordinates. */
export interface FormattedFindMatch extends FindMatch {
	/** Zero-based index of the starting paragraph in document traversal order. */
	readonly paragraphIndex: number;
	/** ID of the starting paragraph. */
	readonly paragraphId: string;
	/** Index of the starting run within the starting paragraph. */
	readonly startRunIndex: number;
	/** Character offset within the starting run. */
	readonly startRunOffset: number;
	/** Zero-based index of the ending paragraph (equals `paragraphIndex` unless crossing `^p`). */
	readonly endParagraphIndex: number;
	/** Index of the ending run within the ending paragraph (inclusive). */
	readonly endRunIndex: number;
	/** Exclusive character offset within the ending run. */
	readonly endRunOffset: number;
	/** Story key when produced by `findDocumentFormattedMatches` (e.g. `'main'`, `'header:0:default'`). */
	readonly storyKey?: string;
}

interface CharPosition {
	readonly kind: 'char' | 'paragraphMark' | 'object';
	readonly paragraphIndex: number;
	readonly paragraph: Paragraph;
	readonly runIndex: number;
	readonly runOffset: number;
	readonly run?: TextRun;
}

interface FlattenedBlocks {
	readonly text: string;
	readonly positions: readonly CharPosition[];
	readonly paragraphs: readonly Paragraph[];
}

function isObjectRun(run: TextRun): boolean {
	return Boolean(
		run.break === 'column' || run.image || run.noteReference || run.equation || run.noteSeparator,
	);
}

/** Flattens a block list (including nested table cells) into Find's searchable string and map. */
export function flattenBlocksForFind(blocks: readonly Block[]): FlattenedBlocks {
	const chars: string[] = [];
	const positions: CharPosition[] = [];
	const paragraphs: Paragraph[] = [];
	let previousContainer: readonly Block[] | null = null;
	let previousParagraphIndex = -1;
	let previousParagraph: Paragraph | null = null;

	const push = (char: string, pos: CharPosition): void => {
		chars.push(char);
		positions.push(pos);
	};

	const walk = (list: readonly Block[]): void => {
		for (const block of list) {
			if (block.type === 'paragraph') {
				const paragraphIndex = paragraphs.length;
				paragraphs.push(block);
				if (previousParagraph !== null) {
					const sameContainer = previousContainer === list;
					const lastRunIndex = Math.max(0, previousParagraph.runs.length - 1);
					const lastRun = previousParagraph.runs[lastRunIndex];
					push(sameContainer ? FIND_PARAGRAPH_MARK : FIND_OBJECT, {
						kind: sameContainer ? 'paragraphMark' : 'object',
						paragraphIndex: previousParagraphIndex,
						paragraph: previousParagraph,
						runIndex: lastRunIndex,
						runOffset: lastRun?.text.length ?? 0,
					});
				}
				block.runs.forEach((run, runIndex) => {
					const posBase = { paragraphIndex, paragraph: block, runIndex, run };
					if (run.break === 'page') {
						push(FIND_PAGE_BREAK, { ...posBase, kind: 'char', runOffset: 0 });
						return;
					}
					if (isObjectRun(run)) {
						push(FIND_OBJECT, { ...posBase, kind: 'object', runOffset: 0 });
						return;
					}
					for (let offset = 0; offset < run.text.length; offset++) {
						const ch = run.text.charAt(offset);
						push(ch === '\n' ? FIND_LINE_BREAK : ch, {
							...posBase,
							kind: 'char',
							runOffset: offset,
						});
					}
				});
				previousContainer = list;
				previousParagraphIndex = paragraphIndex;
				previousParagraph = block;
			} else {
				for (const row of block.rows) {
					for (const cell of row) {
						walk(cell.paragraphs);
					}
				}
			}
		}
	};

	walk(blocks);
	return { text: chars.join(''), positions, paragraphs };
}

function positionMatchesFormat(
	pos: CharPosition,
	criteria: FindFormatCriteria,
	context: FindFormatContext | undefined,
): boolean {
	if (pos.kind === 'object') return false;
	if (pos.kind === 'paragraphMark') {
		return matchesParagraphFormat(pos.paragraph, criteria, context);
	}
	if (!pos.run) return false;
	return matchesFormatCriteria(pos.run, pos.paragraph, criteria, context);
}

function buildMatch(
	flat: FlattenedBlocks,
	start: number,
	end: number,
	text: string,
	groups: readonly string[],
	storyKey?: string,
): FormattedFindMatch | null {
	const first = flat.positions[start];
	const last = flat.positions[end - 1];
	if (!first || !last) return null;
	const endOffset = last.kind === 'char' ? last.runOffset + 1 : last.runOffset;
	return {
		start,
		end,
		text,
		groups,
		paragraphIndex: first.paragraphIndex,
		paragraphId: first.paragraph.id,
		startRunIndex: first.runIndex,
		startRunOffset: first.runOffset,
		endParagraphIndex: last.paragraphIndex,
		endRunIndex: last.runIndex,
		endRunOffset: endOffset,
		...(storyKey !== undefined ? { storyKey } : {}),
	};
}

/** Format-only search: maximal contiguous spans within a paragraph matching `criteria`. */
function findFormatOnlyMatches(
	flat: FlattenedBlocks,
	criteria: FindFormatCriteria,
	context: FindFormatContext | undefined,
	storyKey?: string,
): FormattedFindMatch[] {
	const matches: FormattedFindMatch[] = [];
	let spanStart = -1;
	let spanParagraph = -1;

	const flush = (spanEnd: number): void => {
		if (spanStart < 0 || spanEnd <= spanStart) return;
		const text = flat.text.slice(spanStart, spanEnd);
		const match = buildMatch(flat, spanStart, spanEnd, text, [text], storyKey);
		if (match) matches.push(match);
		spanStart = -1;
		spanParagraph = -1;
	};

	for (let index = 0; index < flat.positions.length; index++) {
		const pos = flat.positions[index];
		if (!pos) continue;
		const ok =
			pos.kind === 'char' &&
			pos.run !== undefined &&
			matchesFormatCriteria(pos.run, pos.paragraph, criteria, context);
		if (ok) {
			if (spanStart < 0) {
				spanStart = index;
				spanParagraph = pos.paragraphIndex;
			} else if (pos.paragraphIndex !== spanParagraph) {
				flush(index);
				spanStart = index;
				spanParagraph = pos.paragraphIndex;
			}
		} else {
			flush(index);
		}
	}
	flush(flat.positions.length);
	return matches;
}

/**
 * Finds every match of `query` and/or `options.format` across `blocks` (including tables):
 * - Empty `query` + non-empty `options.format`: format-only search (maximal contiguous spans).
 * - Non-empty `query` + `options.format`: text search filtered so every character matches `format`.
 * - Non-empty `query` without `format`: plain/wildcard text search with run coordinates.
 */
export function findFormattedMatches(
	blocks: readonly Block[],
	query: string,
	options: FindFormatOptions = {},
	storyKey?: string,
): FormattedFindMatch[] {
	const criteria = options.format;
	const hasCriteria = hasFormatCriteria(criteria);
	if (!query && !hasCriteria) return [];

	const flat = flattenBlocksForFind(blocks);
	if (!flat.text) return [];

	if (!query && criteria) {
		return findFormatOnlyMatches(flat, criteria, options.context, storyKey);
	}

	const textMatches = findMatches(flat.text, query, options);
	const results: FormattedFindMatch[] = [];
	for (const match of textMatches) {
		if (criteria && hasCriteria) {
			let allMatch = true;
			for (let index = match.start; index < match.end; index++) {
				const pos = flat.positions[index];
				if (!pos || !positionMatchesFormat(pos, criteria, options.context)) {
					allMatch = false;
					break;
				}
			}
			if (!allMatch) continue;
		}
		const built = buildMatch(flat, match.start, match.end, match.text, match.groups, storyKey);
		if (built) results.push(built);
	}
	return results;
}

/** Builds a `FindFormatContext` from a `DocumentModel`, overridden by any explicit `override`. */
export function modelFormatContext(
	model: DocumentModel,
	override?: FindFormatContext,
): FindFormatContext {
	const paragraphStyles = override?.paragraphStyles ?? model.paragraphStyles;
	const runStyles = override?.runStyles ?? model.characterStyles;
	const theme = override?.theme ?? model.theme;
	return {
		...(paragraphStyles !== undefined ? { paragraphStyles } : {}),
		...(runStyles !== undefined ? { runStyles } : {}),
		...(theme !== undefined ? { theme } : {}),
	};
}

/** Searches every story of `model` in Word's story order (`findStories`). */
export function findDocumentFormattedMatches(
	model: DocumentModel,
	query: string,
	options: FindFormatOptions = {},
): FormattedFindMatch[] {
	const context = modelFormatContext(model, options.context);
	const mergedOptions: FindFormatOptions = { ...options, context };
	const matches: FormattedFindMatch[] = [];
	for (const story of findStories(model)) {
		const key = findStoryKey(story);
		if (story.kind === 'main') {
			matches.push(...findFormattedMatches(model.blocks, query, mergedOptions, key));
		} else if (story.kind !== 'textBoxes') {
			matches.push(...findFormattedMatches(story.blocks, query, mergedOptions, key));
		}
	}
	return matches;
}
