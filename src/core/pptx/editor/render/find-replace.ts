/**
 * Pure find & replace helpers for PPTX slides, shared by every binding.
 *
 * All functions are immutable; they never mutate their input arrays. No
 * framework imports, so this is safe to use in web workers or server-side code.
 */

import { chartSourceReferenceCounts, findMatchChartSource } from './chart-find-source-guards';
import type { PptxSlide } from '../../index';
import { findSlideTextTargets, isFindBreak, replaceFindRuns } from './find-text-targets';
import { replaceFindTarget } from './find-replace-target';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** Identifies a single match of the find query at segment-level precision. */
export interface FindResult {
	/** Zero-based index of the slide in the slides array. */
	slideIndex: number;
	/** Id of the element that contains the match. */
	elementId: string;
	/** Zero-based index of the text segment within the element's `textSegments`. */
	segmentIndex: number;
	/** Character offset of the match start within the segment text. */
	startOffset: number;
	/** Number of characters matched (equals `query.length`). */
	length: number;
	/** Child indexes inside the selected top-level group, when the match is nested. */
	elementPath?: number[];
	/** Table-cell coordinates, when the match belongs to a table. */
	tableCell?: { rowIndex: number; colIndex: number };
	/** Data-model node id, when the match belongs to SmartArt. */
	smartArtNodeId?: string;
	/** Offset within this object's run stream, used to resolve overlapping replacements. */
	textOffset?: number;
	/** Speaker-note body match, rather than a canvas element. */
	notes?: boolean;
	/** Match in the chart title rather than the shape's own text. */
	chartTitle?: boolean;
	chartSeriesIndex?: number;
	chartCategoryIndex?: number;
	/** Linked or complex chart text is searchable but requires source editing. */
	readOnly?: boolean;
}

/** Options passed to search and replace helpers. */
export interface FindOptions {
	/** When `true`, the search is case-sensitive. Defaults to `false`. */
	matchCase?: boolean;
	/** Search speaker notes as well as editable slide text. */
	includeNotes?: boolean;
	/** Include chart titles. Enabled by default; disable to replace only ordinary slide text. */
	includeChartTitles?: boolean;
	/** Opt-in search of series names and category labels, with transactional source editing. */
	includeChartData?: boolean;
}

/** Result returned by {@link replaceInSlides} and {@link replaceMatch}. */
export interface ReplaceResult {
	/** Updated slides (immutable copy; unchanged slides share the original reference). */
	slides: readonly PptxSlide[];
	/** Number of text replacements actually applied. */
	replacements: number;
}

/** Shared replacement gating for searchable text whose source is not safely editable. */
export function findReplacementAvailability(
	matches: readonly Pick<FindResult, 'readOnly'>[],
	index: number,
) {
	const hasReadOnly = matches.some((match) => match.readOnly);
	return {
		current: Boolean(matches[index] && !matches[index]!.readOnly),
		all: matches.length > 0 && !hasReadOnly,
		hasReadOnly,
	};
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Search all slides for occurrences of `query` at segment-level precision.
 *
 * - Searches shape text, table cells and SmartArt nodes, including nested groups.
 * - Formatting boundaries do not split a match; explicit paragraph/line markers do.
 * - Overlapping matches are counted (search advances by 1 character after each
 *   hit, not by `query.length`).
 * - Returns an empty array when `query` is empty.
 */
export function findInSlides(
	slides: readonly PptxSlide[],
	query: string,
	opts: FindOptions = {},
): FindResult[] {
	if (!query) {
		return [];
	}

	const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const pattern = new RegExp(`(?=(${escaped}))`, opts.matchCase ? 'gu' : 'giu');
	const results: FindResult[] = [];
	const chartSources = chartSourceReferenceCounts(slides);

	for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
		const slide = slides[slideIndex];
		for (const { runs, ...target } of findSlideTextTargets(
			slide,
			opts.includeNotes,
			opts.includeChartTitles,
			opts.includeChartData,
		)) {
			if (
				target.chartTitle ||
				target.chartSeriesIndex !== undefined ||
				target.chartCategoryIndex !== undefined
			) {
				const source = findMatchChartSource(slides, { ...target, slideIndex });
				if (source && (chartSources.get(source) ?? 0) > 1) target.readOnly = true;
			}
			let text = '',
				totalOffset = 0,
				regionOffset = 0;
			let positions: { index: number; offset: number; length: number }[] = [];
			const flush = () => {
				for (const match of text.matchAll(pattern)) {
					const pos = match.index!;
					const run = positions.find(
						(position) => pos >= position.offset && pos < position.offset + position.length,
					);
					if (run)
						results.push({
							...target,
							slideIndex,
							segmentIndex: run.index,
							startOffset: pos - run.offset,
							length: match[1]!.length,
							textOffset: regionOffset + pos,
						});
				}
				text = '';
				positions = [];
			};
			for (const [index, run] of runs.entries()) {
				if (isFindBreak(run)) {
					flush();
					totalOffset++;
					regionOffset = totalOffset;
					continue;
				}
				positions.push({ index, offset: text.length, length: run.text.length });
				text += run.text;
				totalOffset += run.text.length;
			}
			flush();
		}
	}

	return results;
}

/**
 * Apply a set of find-replace substitutions to slides immutably.
 *
 * Replacements inside the same segment are applied in **descending offset**
 * order so that earlier matches are not shifted by later substitutions.
 *
 * Returns the original `slides` reference (no allocation) when `toReplace` is
 * empty.
 */
export function applyFindReplacements(
	slides: readonly PptxSlide[],
	toReplace: readonly FindResult[],
	replacement: string,
): ReplaceResult {
	if (toReplace.length === 0) {
		return { slides, replacements: 0 };
	}
	if (toReplace.some((match) => match.readOnly))
		throw new Error(
			'Linked or complex chart text must be edited in PowerPoint or their source workbook',
		);

	const chartSources = chartSourceReferenceCounts(slides);
	if (
		toReplace.some((match) => {
			if (
				!match.chartTitle &&
				match.chartSeriesIndex === undefined &&
				match.chartCategoryIndex === undefined
			)
				return false;
			const source = findMatchChartSource(slides, match);
			return source && (chartSources.get(source) ?? 0) > 1;
		})
	)
		throw new Error(
			'Text in a shared chart part requires editing in PowerPoint or its source workbook',
		);

	// Group by text object, rather than by run: a match can span several runs.
	const grouped = new Map<string, FindResult[]>();
	for (const match of toReplace) {
		const key = JSON.stringify([
			match.slideIndex,
			match.elementId,
			match.elementPath,
			match.tableCell,
			match.smartArtNodeId,
			match.notes,
			match.chartTitle,
			match.chartSeriesIndex,
			match.chartCategoryIndex,
		]);
		let bucket = grouped.get(key);
		if (!bucket) {
			bucket = [];
			grouped.set(key, bucket);
		}
		bucket.push(match);
	}

	// Work on a shallow-copied array; only affected slides are replaced.
	const nextSlides: PptxSlide[] = slides.slice();
	let totalReplacements = 0;

	for (const [, matches] of grouped) {
		// Sort descending by startOffset so slice indices stay valid.
		const sorted = matches
			.slice()
			.sort((a, b) => b.segmentIndex - a.segmentIndex || b.startOffset - a.startOffset);
		let lastOffset = Infinity;

		for (const match of sorted) {
			const slide = nextSlides[match.slideIndex];
			if (!slide) {
				continue;
			}

			if (match.notes) {
				if (match.textOffset !== undefined && match.textOffset + match.length > lastOffset)
					continue;
				const runs =
					slide.notesSegments?.length &&
					slide.notesSegments.map((run) => run.text).join('') === slide.notes
						? slide.notesSegments
						: [{ text: slide.notes ?? '', style: {} }];
				const updated = replaceFindRuns(
					runs,
					match.segmentIndex,
					match.startOffset,
					match.length,
					replacement,
				);
				if (!updated) continue;
				nextSlides[match.slideIndex] = {
					...slide,
					notes: updated.map((run) => run.text).join(''),
					notesSegments: updated,
				};
				if (match.textOffset !== undefined) lastOffset = match.textOffset;
				totalReplacements++;
				continue;
			}
			const elIdx = slide.elements.findIndex((e) => e.id === match.elementId);
			if (elIdx === -1) {
				continue;
			}

			if (match.textOffset !== undefined && match.textOffset + match.length > lastOffset) continue;
			const element = slide.elements[elIdx]!;
			const updated = replaceFindTarget(element, match, replacement);
			if (updated === element) continue;
			if (match.textOffset !== undefined) lastOffset = match.textOffset;
			const nextElements = slide.elements.slice();
			nextElements[elIdx] = updated;

			nextSlides[match.slideIndex] = { ...slide, elements: nextElements };
			totalReplacements++;
		}
	}

	return { slides: nextSlides, replacements: totalReplacements };
}

/**
 * Replace a single match (identified by its position in `allResults`) and
 * return the updated slides together with the count of replacements applied
 * (always 0 or 1).
 *
 * This is a convenience wrapper around {@link applyFindReplacements} for the
 * "Replace current" action.
 */
export function replaceMatch(
	slides: readonly PptxSlide[],
	allResults: readonly FindResult[],
	matchIndex: number,
	replacement: string,
): ReplaceResult {
	const match = allResults[matchIndex];
	if (!match) {
		return { slides, replacements: 0 };
	}
	return applyFindReplacements(slides, [match], replacement);
}

/**
 * Replace **all** current matches and return the updated slides.
 *
 * This is a convenience wrapper around {@link applyFindReplacements} for the
 * "Replace All" action.
 */
export function replaceInSlides(
	slides: readonly PptxSlide[],
	query: string,
	replacement: string,
	opts: FindOptions = {},
): ReplaceResult {
	const results = findInSlides(slides, query, opts);
	return applyFindReplacements(slides, results, replacement);
}
