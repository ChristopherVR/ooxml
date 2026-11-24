// Formatted Replace and Replace All over DOCX blocks and DocumentModel stories. Applies
// `ReplaceFormatPatch` changes without losing untouched run or paragraph properties, and handles
// Word's format-only replacement rule (empty replacement text + replaceFormat preserves text).
import {
	applyParagraphFormatPatch,
	applyRunFormatPatch,
	hasReplaceFormat,
	type FindFormatContext,
	type ReplaceFormatPatch,
} from './find-format-criteria';
import {
	findFormattedMatches,
	type FindFormatOptions,
	type FormattedFindMatch,
} from './find-format';
import { replacementParts, type ReplacementPart } from './find-replacement';
import type { Block, Paragraph, TextRun } from './model';

/** Options for formatted Replace / Replace All over DOCX blocks or a `DocumentModel`. */
export interface ReplaceFormatOptions extends FindFormatOptions {
	/** Formatting changes to apply to replaced runs and paragraphs. */
	readonly replaceFormat?: ReplaceFormatPatch;
	/** When set, replaces only the single match at this zero-based index; otherwise replaces all. */
	readonly matchIndex?: number;
}

export interface ReplaceFormatResult {
	readonly blocks: Block[];
	readonly replacedCount: number;
}

export function cloneBlocks(blocks: readonly Block[]): Block[] {
	return blocks.map((block): Block => {
		if (block.type === 'paragraph') {
			return { ...block, runs: block.runs.map((run) => ({ ...run })) };
		}
		return {
			...block,
			rows: block.rows.map((row) =>
				row.map((cell) => ({
					...cell,
					paragraphs: cloneBlocks(cell.paragraphs) as Paragraph[],
				})),
			),
		};
	});
}

interface ParagraphLocation {
	readonly container: Block[];
	readonly index: number;
}

function collectParagraphLocations(
	blocks: Block[],
	out: ParagraphLocation[] = [],
): ParagraphLocation[] {
	for (let index = 0; index < blocks.length; index++) {
		const block = blocks[index];
		if (!block) continue;
		if (block.type === 'paragraph') {
			out.push({ container: blocks, index });
		} else {
			for (const row of block.rows) {
				for (const cell of row) {
					collectParagraphLocations(cell.paragraphs, out);
				}
			}
		}
	}
	return out;
}

function partToRuns(parts: readonly ReplacementPart[], templateRun: TextRun): TextRun[][] {
	const groups: TextRun[][] = [[]];
	const baseRun = (): TextRun => {
		const copy: TextRun = { ...templateRun, text: '' };
		delete copy.break;
		return copy;
	};
	const pushText = (value: string): void => {
		const current = groups[groups.length - 1];
		if (!current) return;
		const last = current[current.length - 1];
		if (last && !last.break) last.text += value;
		else current.push({ ...baseRun(), text: value });
	};
	for (const part of parts) {
		switch (part.kind) {
			case 'text':
				pushText(part.text);
				break;
			case 'paragraph':
				groups.push([]);
				break;
			case 'tab':
				pushText('\t');
				break;
			case 'lineBreak':
				pushText('\n');
				break;
			case 'pageBreak':
				groups[groups.length - 1]?.push({ ...baseRun(), text: '', break: 'page' });
				break;
			case 'optionalHyphen':
				pushText('\u00AD');
				break;
			case 'nonBreakingHyphen':
				pushText('\u2011');
				break;
			case 'nonBreakingSpace':
				pushText('\u00A0');
				break;
		}
	}
	return groups;
}

/** Splits runs of a single paragraph around `[startRunIndex:startRunOffset, endRunIndex:endRunOffset)`. */
function splitParagraphRuns(
	runs: readonly TextRun[],
	startRunIndex: number,
	startRunOffset: number,
	endRunIndex: number,
	endRunOffset: number,
): { before: TextRun[]; middle: TextRun[]; after: TextRun[] } {
	const before: TextRun[] = [];
	const middle: TextRun[] = [];
	const after: TextRun[] = [];

	for (let i = 0; i < runs.length; i++) {
		const run = runs[i];
		if (!run) continue;
		if (i < startRunIndex) {
			before.push(run);
		} else if (i > endRunIndex) {
			after.push(run);
		} else if (i === startRunIndex && i === endRunIndex) {
			if (startRunOffset > 0) before.push({ ...run, text: run.text.slice(0, startRunOffset) });
			const midText = run.text.slice(startRunOffset, endRunOffset);
			if (midText || run.break === 'page') middle.push({ ...run, text: midText });
			if (endRunOffset < run.text.length)
				after.push({ ...run, text: run.text.slice(endRunOffset) });
		} else if (i === startRunIndex) {
			if (startRunOffset > 0) before.push({ ...run, text: run.text.slice(0, startRunOffset) });
			const midText = run.text.slice(startRunOffset);
			if (midText || run.break === 'page') middle.push({ ...run, text: midText });
		} else if (i === endRunIndex) {
			const midText = run.text.slice(0, endRunOffset);
			if (midText || run.break === 'page') middle.push({ ...run, text: midText });
			if (endRunOffset < run.text.length)
				after.push({ ...run, text: run.text.slice(endRunOffset) });
		} else {
			middle.push(run);
		}
	}
	return { before, middle, after };
}

function applySingleReplacement(
	locations: ParagraphLocation[],
	match: FormattedFindMatch,
	replacement: string,
	options: ReplaceFormatOptions,
	context: FindFormatContext | undefined,
): boolean {
	const startLoc = locations[match.paragraphIndex];
	const endLoc = locations[match.endParagraphIndex];
	if (!startLoc || !endLoc) return false;
	const startBlock = startLoc.container[startLoc.index];
	const endBlock = endLoc.container[endLoc.index];
	if (startBlock?.type !== 'paragraph' || endBlock?.type !== 'paragraph') return false;

	const patch = options.replaceFormat;
	const keepTextWithFormat = replacement === '' && hasReplaceFormat(patch);

	// Format-only replacement (empty replacement string + replaceFormat): keep matched text and
	// overlay run/paragraph formatting on every touched run slice and paragraph.
	if (keepTextWithFormat && patch) {
		for (let pIdx = match.paragraphIndex; pIdx <= match.endParagraphIndex; pIdx++) {
			const loc = locations[pIdx];
			if (!loc) continue;
			const block = loc.container[loc.index];
			if (block?.type !== 'paragraph') continue;
			const rStart = pIdx === match.paragraphIndex ? match.startRunIndex : 0;
			const oStart = pIdx === match.paragraphIndex ? match.startRunOffset : 0;
			const rEnd =
				pIdx === match.endParagraphIndex ? match.endRunIndex : Math.max(0, block.runs.length - 1);
			const oEnd =
				pIdx === match.endParagraphIndex
					? match.endRunOffset
					: (block.runs[rEnd]?.text.length ?? 0);
			const { before, middle, after } = splitParagraphRuns(block.runs, rStart, oStart, rEnd, oEnd);
			const patchedMiddle = middle.map((run) => applyRunFormatPatch(run, patch, context));
			const updated = applyParagraphFormatPatch(
				{ ...block, runs: [...before, ...patchedMiddle, ...after] },
				patch,
				context,
			);
			loc.container[loc.index] = updated;
		}
		return true;
	}

	const startSplit = splitParagraphRuns(
		startBlock.runs,
		match.startRunIndex,
		match.startRunOffset,
		match.paragraphIndex === match.endParagraphIndex
			? match.endRunIndex
			: Math.max(0, startBlock.runs.length - 1),
		match.paragraphIndex === match.endParagraphIndex
			? match.endRunOffset
			: (startBlock.runs[Math.max(0, startBlock.runs.length - 1)]?.text.length ?? 0),
	);
	const afterRuns =
		match.paragraphIndex === match.endParagraphIndex
			? startSplit.after
			: splitParagraphRuns(endBlock.runs, 0, 0, match.endRunIndex, match.endRunOffset).after;

	const templateRun: TextRun = startSplit.middle[0] ??
		startBlock.runs[match.startRunIndex] ?? { text: '' };
	const patchedTemplate = patch ? applyRunFormatPatch(templateRun, patch, context) : templateRun;
	const parts = replacement ? replacementParts(replacement, match, Boolean(options.wildcards)) : [];
	const runGroups = parts.length > 0 ? partToRuns(parts, patchedTemplate) : [[]];

	const baseParagraph = patch ? applyParagraphFormatPatch(startBlock, patch, context) : startBlock;
	const newParagraphs: Paragraph[] = runGroups.map((runs, groupIndex) => {
		const leading = groupIndex === 0 ? startSplit.before : [];
		const trailing = groupIndex === runGroups.length - 1 ? afterRuns : [];
		return {
			...baseParagraph,
			id: groupIndex === 0 ? startBlock.id : `${startBlock.id}-r${groupIndex}`,
			runs: [...leading, ...runs, ...trailing],
		};
	});

	// Replace the spanned paragraphs in the same container (or single paragraph).
	const removeCount =
		startLoc.container === endLoc.container ? endLoc.index - startLoc.index + 1 : 1;
	startLoc.container.splice(startLoc.index, removeCount, ...newParagraphs);
	return true;
}

/**
 * Replaces matched text and/or applies `options.replaceFormat` across `blocks`, returning a new
 * block array and the number of matches replaced. Never mutates the input blocks.
 */
export function replaceFormattedMatches(
	blocks: readonly Block[],
	query: string,
	replacement: string,
	options: ReplaceFormatOptions = {},
): ReplaceFormatResult {
	const matches = findFormattedMatches(blocks, query, options);
	if (matches.length === 0) return { blocks: cloneBlocks(blocks), replacedCount: 0 };

	let targets: readonly FormattedFindMatch[];
	if (options.matchIndex !== undefined) {
		const single = matches[options.matchIndex];
		if (!single) return { blocks: cloneBlocks(blocks), replacedCount: 0 };
		targets = [single];
	} else {
		targets = matches;
	}

	const cloned = cloneBlocks(blocks);
	let replacedCount = 0;
	for (let i = targets.length - 1; i >= 0; i--) {
		const match = targets[i];
		if (!match) continue;
		const locations = collectParagraphLocations(cloned);
		if (applySingleReplacement(locations, match, replacement, options, options.context)) {
			replacedCount++;
		}
	}
	return { blocks: cloned, replacedCount };
}
