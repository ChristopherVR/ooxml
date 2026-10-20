// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// STYLEREF fields in headers and footers, page by page, as Word lays them out (ECMA-376 Part 1,
// 17.16.5, the search for a printed document; MS-OE376 2.1.530 c and f say Word uses it on screen
// too and how `\l` reads).
import type { DocumentModel, HeaderFooterContent } from '../model';
import {
	styleRefSource,
	type StyleRefPortion,
	type StyleRefSource,
	type StyleReference,
} from '../style-ref';
import { withStyleRefResults } from '../style-ref-fields';
import type { LayoutBlockBox, LayoutPageBox, LayoutParagraphBox } from './result';
import { reviewSourceLength } from './review-display';

/** STYLEREF results of the pages of one layout. */
export interface PageStyleRefs {
	/** The result a header or footer STYLEREF shows on `pages[page]`, or undefined to keep its own. */
	result(page: number, reference: StyleReference): string | undefined;
	/** A header or footer as it shows on `pages[page]`: its STYLEREF results filled in. */
	content(content: HeaderFooterContent, page: number): HeaderFooterContent;
}

/** The parts of a paragraph a page shows: its lines' source ranges, in the paragraph's positions. */
type Shown = [start: number, end: number][];

/**
 * The main document's paragraphs each page shows, as `source` indexes top to bottom, with the
 * parts of each that are on the page (a paragraph can run over several pages).
 */
function paragraphsByPage(
	pages: readonly LayoutPageBox[],
	source: StyleRefSource,
): Map<number, Shown>[] {
	return pages.map((page) => {
		const found = new Map<number, Shown>();
		const add = (box: LayoutParagraphBox) => {
			const at = source.index.get(box.blockId);
			if (at === undefined) return;
			const shown = found.get(at) ?? [];
			for (const line of box.lines) shown.push([line.sourceStart, line.sourceEnd]);
			found.set(at, shown);
		};
		const visit = (blocks: readonly LayoutBlockBox[]) => {
			for (const block of blocks)
				if (block.kind === 'paragraph') add(block);
				else
					for (const row of block.rows)
						// A header row repeated on a later page is not text of that page.
						if (!row.repeated) for (const cell of row.cells) for (const box of cell) add(box);
		};
		for (const column of page.columns) visit(column.blocks);
		return found;
	});
}

/** The last entry of an ascending list below `value`, or the first above it, by binary search. */
function nearest(sorted: readonly number[], value: number, direction: 'before' | 'after') {
	let low = 0;
	let high = sorted.length;
	while (low < high) {
		const middle = (low + high) >> 1;
		if (direction === 'before' ? sorted[middle]! < value : sorted[middle]! <= value)
			low = middle + 1;
		else high = middle;
	}
	return direction === 'before' ? sorted[low - 1] : sorted[low];
}

/**
 * Evaluates header and footer STYLEREF fields for each page of a layout of `model`. On a page,
 * the field takes the first text in the style from the top (with `\l`, the last from the bottom);
 * when the page has none, the nearest before the page, then the nearest after it. Only the text
 * a page shows counts for it, so a paragraph that runs over two pages gives each page its own.
 */
export function pageStyleRefs(
	model: DocumentModel,
	pages: readonly LayoutPageBox[],
): PageStyleRefs {
	let source: StyleRefSource | undefined;
	let onPage: Map<number, Shown>[] | undefined;
	let bounds: { first: number; last: number }[] | undefined;
	const matches = new Map<string, { portions: Map<number, StyleRefPortion[]>; sorted: number[] }>();
	const offsets = new Map<number, number[]>();
	const read = () => {
		if (!source) {
			source = styleRefSource(model);
			onPage = paragraphsByPage(pages, source);
			let previous = -1;
			bounds = onPage.map((here) => {
				if (!here.size) return { first: previous + 1, last: previous };
				const range = { first: Math.min(...here.keys()), last: Math.max(...here.keys()) };
				previous = range.last;
				return range;
			});
		}
		return source;
	};
	/** Where a portion starts and ends in its paragraph's positions, as layout counts them. */
	const span = (at: number, portion: StyleRefPortion): [number, number] => {
		let starts = offsets.get(at);
		if (!starts) {
			starts = [0];
			for (const run of source!.paragraphs[at]!.runs)
				starts.push(starts.at(-1)! + reviewSourceLength(run));
			offsets.set(at, starts);
		}
		return [starts[portion.from]!, starts[portion.to + 1]!];
	};
	const result = (page: number, reference: StyleReference): string | undefined => {
		const document = read();
		const target = document.target(reference.style);
		if (!target) return undefined;
		const key = JSON.stringify(reference);
		let found = matches.get(key);
		if (!found) {
			const portions = new Map<number, StyleRefPortion[]>();
			for (let at = 0; at < document.paragraphs.length; at++) {
				const results = document.portions(at, reference, target);
				if (results.length) portions.set(at, results);
			}
			found = { portions, sorted: [...portions.keys()] };
			matches.set(key, found);
		}
		const here = onPage![page] ?? new Map<number, Shown>();
		const inParagraph = (at: number, keep: (start: number, end: number) => boolean) =>
			(found.portions.get(at) ?? []).filter((portion) => keep(...span(at, portion)));
		const order = [...here.keys()];
		for (const at of reference.last ? order.reverse() : order) {
			const shown = here.get(at)!;
			const texts = inParagraph(at, (start, end) =>
				shown.some(([from, to]) => start < to && end > from),
			);
			if (texts.length) return (reference.last ? texts.at(-1) : texts[0])!.text;
		}
		const range = bounds![page] ?? { first: 0, last: -1 };
		// The page's first paragraph may have begun on an earlier page, and its last may go on.
		const first = here.get(range.first);
		if (first) {
			const top = Math.min(...first.map(([from]) => from));
			const earlier = inParagraph(range.first, (_start, end) => end <= top);
			if (earlier.length) return earlier.at(-1)!.text;
		}
		const before = nearest(found.sorted, range.first, 'before');
		if (before !== undefined) return found.portions.get(before)!.at(-1)!.text;
		const last = here.get(range.last);
		if (last) {
			const bottom = Math.max(...last.map(([, to]) => to));
			const later = inParagraph(range.last, (start) => start >= bottom);
			if (later.length) return later[0]!.text;
		}
		const after = nearest(found.sorted, range.last, 'after');
		return after === undefined ? undefined : found.portions.get(after)![0]!.text;
	};
	return {
		result,
		content(content, page) {
			const blocks = withStyleRefResults(content.blocks, (reference) => result(page, reference));
			return blocks === content.blocks ? content : { ...content, blocks };
		},
	};
}
