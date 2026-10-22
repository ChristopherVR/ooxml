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
 * parts of each that are on the page (a paragraph can run over several pages). A box without
 * lines does not say which part of its paragraph the page holds, so all of it counts.
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
			if (!box.lines.length) shown.push([0, Number.POSITIVE_INFINITY]);
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

/** Whether `shown` holds the part `start` to `end` of a paragraph (an empty part, its place). */
const shows = (shown: Shown, start: number, end: number) =>
	shown.some(([from, to]) =>
		start === end ? from <= start && start <= to : start < to && end > from,
	);

/** A match and the first and last page that show it. */
interface Placed {
	text: string | undefined;
	first: number;
	last: number;
}

/** The matches of one reference, in reading order, and the nearest of them around each page. */
interface Matches {
	portions: Map<number, StyleRefPortion[]>;
	placed: Placed[];
	/** For each page, the last match in reading order that only pages before it show, or -1. */
	before: number[];
	/** For each page, the first match in reading order that only pages after it show, or -1. */
	after: number[];
}

/**
 * Evaluates header and footer STYLEREF fields for each page of a layout of `model`. On a page,
 * the field takes the first text in the style from the top (with `\l`, the last from the bottom);
 * when the page has none, the nearest before the page, then the nearest after it. Only the text
 * a page shows counts for it, so a paragraph that runs over two pages gives each page its own,
 * and text is before or after a page by the pages that show it: in a table row that runs onto
 * the next page, a cell's text there is after the first page, though it comes before the text of
 * the next cell in reading order.
 */
export function pageStyleRefs(
	model: DocumentModel,
	pages: readonly LayoutPageBox[],
): PageStyleRefs {
	let source: StyleRefSource | undefined;
	let onPage: Map<number, Shown>[] | undefined;
	/** The pages that show each paragraph, first to last. */
	let pagesOf: Map<number, number[]> | undefined;
	const matches = new Map<string, Matches>();
	const offsets = new Map<number, number[]>();
	const read = () => {
		if (!source) {
			source = styleRefSource(model);
			onPage = paragraphsByPage(pages, source);
			pagesOf = new Map();
			onPage.forEach((here, page) => {
				for (const at of here.keys()) {
					const list = pagesOf!.get(at);
					if (list) list.push(page);
					else pagesOf!.set(at, [page]);
				}
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
	/**
	 * Places the matches of one reference. A paragraph no page shows sits between its neighbours:
	 * after the last page of the shown paragraph before it, before the first page of the one after.
	 */
	const place = (portions: Map<number, StyleRefPortion[]>): Matches => {
		const count = source!.paragraphs.length;
		const lastBefore: number[] = [];
		for (let at = 0, last = -1; at < count; at++) {
			lastBefore[at] = last;
			last = pagesOf!.get(at)?.at(-1) ?? last;
		}
		const firstAfter: number[] = [];
		for (let at = count - 1, first = pages.length; at >= 0; at--) {
			firstAfter[at] = first;
			first = pagesOf!.get(at)?.[0] ?? first;
		}
		const placed: Placed[] = [];
		for (const [at, found] of portions) {
			const showing = pagesOf!.get(at);
			for (const portion of found) {
				if (!showing) {
					placed.push({ text: portion.text, first: firstAfter[at]!, last: lastBefore[at]! });
					continue;
				}
				const [start, end] = span(at, portion);
				const on = showing.filter((page) => shows(onPage![page]!.get(at)!, start, end));
				const list = on.length ? on : showing;
				placed.push({ text: portion.text, first: list[0]!, last: list.at(-1)! });
			}
		}
		// A match is before every page after its last one, and after every page before its first.
		const before = new Array<number>(pages.length).fill(-1);
		const after = new Array<number>(pages.length).fill(-1);
		placed.forEach((item, index) => {
			if (item.last + 1 < pages.length) before[Math.max(0, item.last + 1)] = index;
		});
		for (let index = placed.length - 1; index >= 0; index--) {
			const page = placed[index]!.first - 1;
			if (page >= 0) after[Math.min(page, pages.length - 1)] = index;
		}
		for (let page = 1; page < pages.length; page++)
			before[page] = Math.max(before[page]!, before[page - 1]!);
		for (let page = pages.length - 2; page >= 0; page--)
			if (after[page]! < 0 || (after[page + 1]! >= 0 && after[page + 1]! < after[page]!))
				after[page] = after[page + 1]!;
		return { portions, placed, before, after };
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
			found = place(portions);
			matches.set(key, found);
		}
		const here = onPage![page] ?? new Map<number, Shown>();
		const order = [...here.keys()];
		for (const at of reference.last ? order.reverse() : order) {
			const shown = here.get(at)!;
			// An empty match (an empty paragraph in the style) is on the page that shows its line.
			const texts = (found.portions.get(at) ?? []).filter((portion) =>
				shows(shown, ...span(at, portion)),
			);
			if (texts.length) return (reference.last ? texts.at(-1) : texts[0])!.text;
		}
		const nearest =
			(found.before[page] ?? -1) >= 0 ? found.before[page]! : (found.after[page] ?? -1);
		return nearest >= 0 ? found.placed[nearest]!.text : undefined;
	};
	return {
		result,
		content(content, page) {
			const blocks = withStyleRefResults(content.blocks, (reference) => result(page, reference));
			return blocks === content.blocks ? content : { ...content, blocks };
		},
	};
}
