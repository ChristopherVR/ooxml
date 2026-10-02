// Selection overlay of one quadrant, in plane coordinates: shaded ranges with the active cell cut
// out, the green outline and fill handle, the copy marquee (marching ants), the coloured
// reference boxes of the formula being edited and the fill-drag preview.
import type { CellRange, GridMetrics } from '@christophervr/xlsx-core';
import type { Selection } from '../context.js';
import { h, place } from './dom.js';
import { isSingleCell } from './selection-ops.js';

export interface OverlayState {
	selection: Selection | undefined;
	/** The cell range the active cell covers (merge-sized). */
	activeRange: CellRange | undefined;
	copyRange?: CellRange | undefined;
	references?: { range: CellRange; color: string }[] | undefined;
	fillPreview?: CellRange | undefined;
	showHandle: boolean;
	/** Last row/column worth painting (whole-column ranges are clipped to it). */
	maxRow: number;
	maxCol: number;
}

interface Rect {
	x: number;
	y: number;
	w: number;
	h: number;
}

function rectOf(metrics: GridMetrics, range: CellRange, maxRow: number, maxCol: number): Rect {
	const endRow = Math.min(range.end.row, maxRow);
	const endCol = Math.min(range.end.col, maxCol);
	const x = metrics.colLeft(range.start.col);
	const y = metrics.rowTop(range.start.row);
	return { x, y, w: metrics.colLeft(endCol + 1) - x, h: metrics.rowTop(endRow + 1) - y };
}

export class SelectionLayer {
	readonly element: HTMLDivElement;
	readonly #doc: Document;
	readonly #shades: HTMLDivElement[] = [];
	readonly #refs: HTMLDivElement[] = [];
	readonly #outline: HTMLDivElement;
	readonly #active: HTMLDivElement;
	readonly #handle: HTMLDivElement;
	readonly #ants: HTMLDivElement;
	readonly #preview: HTMLDivElement;

	constructor(doc: Document) {
		this.#doc = doc;
		this.element = h(doc, 'div', 'xg-overlay');
		this.#outline = h(doc, 'div', 'xg-outline');
		this.#active = h(doc, 'div', 'xg-active');
		this.#handle = h(doc, 'div', 'xg-handle', { 'data-handle': 'fill' });
		this.#ants = h(doc, 'div', 'xg-ants');
		this.#preview = h(doc, 'div', 'xg-fill-preview');
		this.element.append(this.#outline, this.#active, this.#ants, this.#preview, this.#handle);
		for (const node of [this.#outline, this.#active, this.#ants, this.#preview, this.#handle])
			node.hidden = true;
	}

	#pooled(list: HTMLDivElement[], index: number, className: string): HTMLDivElement {
		let node = list[index];
		if (!node) {
			node = h(this.#doc, 'div', className);
			list.push(node);
			this.element.insertBefore(node, this.#outline);
		}
		node.hidden = false;
		return node;
	}

	render(metrics: GridMetrics, state: OverlayState): void {
		const { selection, maxRow, maxCol } = state;
		const ranges = selection?.ranges ?? [];
		const active = state.activeRange
			? rectOf(metrics, state.activeRange, maxRow, maxCol)
			: undefined;
		let i = 0;
		const multiCell = ranges.length > 1 || !ranges[0] || !isSingleCell(ranges[0]);
		if (multiCell)
			for (const range of ranges) {
				const r = rectOf(metrics, range, maxRow, maxCol);
				const node = this.#pooled(this.#shades, i++, 'xg-shade');
				place(node, r.x, r.y, r.w, r.h);
				node.style.clipPath = active ? cutout(r, active) : '';
			}
		for (let j = i; j < this.#shades.length; j++) {
			const node = this.#shades[j];
			if (node) node.hidden = true;
		}
		const last = ranges[ranges.length - 1];
		const single = ranges.length === 1 && last;
		this.#outline.hidden = !single;
		let handleAt: Rect | undefined;
		if (single) {
			const r = rectOf(metrics, last, maxRow, maxCol);
			place(this.#outline, r.x - 1, r.y - 1, r.w + 1, r.h + 1);
			handleAt = r;
		}
		this.#active.hidden = !active || Boolean(single && last && isSingleCell(last));
		if (active && !this.#active.hidden)
			place(this.#active, active.x - 1, active.y - 1, active.w + 1, active.h + 1);
		if (single && last && isSingleCell(last) && active) {
			place(this.#outline, active.x - 1, active.y - 1, active.w + 1, active.h + 1);
			handleAt = active;
		}
		this.#handle.hidden = !state.showHandle || !handleAt;
		if (handleAt && state.showHandle)
			place(this.#handle, handleAt.x + handleAt.w - 4, handleAt.y + handleAt.h - 4, 7, 7);
		this.#ants.hidden = !state.copyRange;
		if (state.copyRange) {
			const r = rectOf(metrics, state.copyRange, maxRow, maxCol);
			place(this.#ants, r.x - 1, r.y - 1, r.w + 1, r.h + 1);
		}
		this.#preview.hidden = !state.fillPreview;
		if (state.fillPreview) {
			const r = rectOf(metrics, state.fillPreview, maxRow, maxCol);
			place(this.#preview, r.x - 1, r.y - 1, r.w + 1, r.h + 1);
		}
		let k = 0;
		for (const ref of state.references ?? []) {
			const r = rectOf(metrics, ref.range, maxRow, maxCol);
			const node = this.#pooled(this.#refs, k++, 'xg-refbox');
			place(node, r.x - 1, r.y - 1, r.w + 1, r.h + 1);
			node.style.borderColor = ref.color;
			node.style.backgroundColor = `color-mix(in srgb, ${ref.color} 12%, transparent)`;
		}
		for (let j = k; j < this.#refs.length; j++) {
			const node = this.#refs[j];
			if (node) node.hidden = true;
		}
	}
}

/** clip-path polygon showing `outer` minus `inner` (both in plane pixels). */
export function cutout(outer: Rect, inner: Rect): string {
	const x1 = inner.x - outer.x;
	const y1 = inner.y - outer.y;
	const x2 = x1 + inner.w;
	const y2 = y1 + inner.h;
	if (x2 <= 0 || y2 <= 0 || x1 >= outer.w || y1 >= outer.h) return '';
	return `polygon(evenodd, 0 0, ${outer.w}px 0, ${outer.w}px ${outer.h}px, 0 ${outer.h}px, 0 0, ${x1}px ${y1}px, ${x2}px ${y1}px, ${x2}px ${y2}px, ${x1}px ${y2}px, ${x1}px ${y1}px)`;
}
