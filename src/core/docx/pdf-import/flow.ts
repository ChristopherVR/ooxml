/**
 * A page's content in reading order: single-column runs and column regions top to bottom, each
 * column read in turn, with the borderless tables, ruled tables and pictures each one holds.
 */
import { splitColumns, type ColumnBand } from './columns';
import type { Line } from './lines';
import type { PictureBlock } from './package';
import {
	buildParagraphs,
	continuesOnNextPage,
	type ParagraphBlock,
	type TextArea,
} from './paragraphs';
import { unruledTables, type TableBlock } from './tables';

export type FlowItem = ParagraphBlock | PictureBlock | TableBlock;

export interface ColumnLayout {
	count: number;
	/** Space between the columns, points. */
	space: number;
	/** Column widths when they differ, points. */
	widths?: number[];
}

export interface PageFlow {
	items: FlowItem[];
	/** The column layout of each item (undefined for one column). */
	layouts: (ColumnLayout | undefined)[];
}

type Placed = (PictureBlock & { x: number }) | TableBlock;

export const itemTop = (item: FlowItem) =>
	item.kind === 'paragraph' ? item.lines[0]!.top : item.top;
export const itemBottom = (item: FlowItem) => {
	if (item.kind === 'picture') return item.top + item.height;
	if (item.kind === 'table') return item.bottom;
	return item.lines[item.lines.length - 1]!.bottom;
};
const middleX = (item: Placed) =>
	item.kind === 'picture'
		? item.x + item.width / 2
		: item.left + item.columns.reduce((sum, width) => sum + width, 0) / 2;

function flowOf(page: number, lines: Line[], area: TextArea, bodySize: number, placed: Placed[]) {
	const { tables, lines: rest } = unruledTables(page, lines, bodySize);
	const items: FlowItem[] = [...buildParagraphs(rest, area, bodySize), ...tables, ...placed];
	for (const item of items) if (item.kind === 'table') item.indent = item.left - area.left;
	return items.sort((a, b) => itemTop(a) - itemTop(b));
}

function layoutOf(band: ColumnBand): ColumnLayout {
	const widths = band.bounds.map(([left, right]) => right - left);
	const equal = widths.every((width) => Math.abs(width - widths[0]!) < widths[0]! * 0.05);
	return {
		count: band.columns.length,
		space: band.space,
		...(equal ? {} : { widths }),
	};
}

export function pageFlow(
	page: {
		index?: number;
		lines: Line[];
		pictures: (PictureBlock & { x: number })[];
		tables?: TableBlock[];
	},
	area: TextArea,
	bodySize: number,
): PageFlow {
	const index = page.lines[0]?.page ?? page.index ?? 0;
	const bands = splitColumns(page.lines, area);
	const placed: Placed[] = [...page.pictures, ...(page.tables ?? [])];
	const columnBands = bands.filter((band): band is ColumnBand => band.kind === 'columns');
	const regionOf = (item: Placed) =>
		columnBands.find((band) => item.top >= band.top - 2 && item.top <= band.bottom);
	const flow: PageFlow = { items: [], layouts: [] };
	// Paragraphs measure their space from the line before; the first one of a band from the band above.
	let bandBottom: number | undefined;
	const spaceFromAbove = (items: FlowItem[]) => {
		const first = items[0];
		if (first?.kind !== 'paragraph' || bandBottom === undefined) return;
		const line = first.lines[0]!;
		const pitch = line.size * 1.2 * first.lineSpacing;
		const space = itemTop(first) - bandBottom - (pitch - 1.05 * line.size);
		first.spaceBefore = space < 2 ? 0 : space;
	};
	const finishBand = (items: FlowItem[]) => {
		if (items.length) bandBottom = Math.max(...items.map(itemBottom));
	};
	bands.forEach((band, at) => {
		if (band.kind === 'single') {
			const above = bands
				.slice(0, at)
				.filter((other) => other.kind === 'columns')
				.pop();
			const below = bands.slice(at + 1).find((other) => other.kind === 'columns');
			const mine = placed.filter(
				(item) =>
					!regionOf(item) &&
					(!above || item.top > (above as ColumnBand).bottom) &&
					(!below || item.top < (below as ColumnBand).top),
			);
			const items = flowOf(index, band.lines, area, bodySize, mine);
			spaceFromAbove(items);
			finishBand(items);
			flow.items.push(...items);
			flow.layouts.push(...items.map(() => undefined));
			return;
		}
		const layout = layoutOf(band);
		const cuts = band.bounds.slice(1).map(([left], k) => (left + band.bounds[k]![1]) / 2);
		let previous: FlowItem[] = [];
		const all: FlowItem[] = [];
		band.columns.forEach((lines, k) => {
			const [left, right] = band.bounds[k]!;
			const mine = placed.filter((item) => {
				if (regionOf(item) !== band) return false;
				const x = middleX(item);
				return (k === 0 || x >= cuts[k - 1]!) && (k === cuts.length || x < cuts[k]!);
			});
			const items = flowOf(index, lines, { left, right }, bodySize, mine);
			// A paragraph the column break cut goes on as one paragraph.
			const last = previous[previous.length - 1];
			const first = items[0];
			if (
				last?.kind === 'paragraph' &&
				first?.kind === 'paragraph' &&
				continuesOnNextPage(last, first)
			) {
				const tail = last.fragments[last.fragments.length - 1];
				if (tail) tail.text = `${tail.text} `;
				last.fragments.push(...first.fragments);
				last.lines.push(...first.lines);
				items.shift();
			}
			if (k === 0) spaceFromAbove(items);
			flow.items.push(...items);
			flow.layouts.push(...items.map(() => layout));
			all.push(...items);
			previous = items;
		});
		finishBand(all);
	});
	// Pictures and tables outside every band (a page with no text) still go in.
	const missing = placed.filter((item) => !flow.items.includes(item));
	for (const item of missing) {
		const position = flow.items.findIndex((other) => itemTop(other) > item.top);
		const at = position < 0 ? flow.items.length : position;
		if (item.kind === 'table') item.indent = item.left - area.left;
		flow.items.splice(at, 0, item);
		flow.layouts.splice(at, 0, flow.layouts[at - 1]);
	}
	return flow;
}
