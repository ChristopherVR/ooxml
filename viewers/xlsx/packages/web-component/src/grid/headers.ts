// Row and column headers of one pane (frozen or scrolling part), keyed by index and recycled.
// Headers of selected rows/columns are highlighted ('part'), fully selected ones darker ('full').
import { columnLabel, type GridMetrics } from '@christophervr/xlsx-core';
import { h } from './dom.js';

export type HeaderAxis = 'row' | 'col';
export type HeaderState = 'none' | 'part' | 'full';

interface HeaderNode extends HTMLDivElement {
	xgSig?: string | undefined;
}

export class HeaderLayer {
	readonly element: HTMLDivElement;
	readonly #doc: Document;
	readonly #axis: HeaderAxis;
	readonly #nodes = new Map<number, HeaderNode>();
	readonly #pool: HeaderNode[] = [];

	constructor(doc: Document, axis: HeaderAxis) {
		this.#doc = doc;
		this.#axis = axis;
		this.element = h(doc, 'div', `xg-hdr-plane xg-hdr-${axis}`);
	}

	render(
		metrics: GridMetrics,
		indices: readonly number[],
		state: (index: number) => HeaderState,
		fontPx: number,
	): void {
		const keep = new Set<number>();
		const row = this.#axis === 'row';
		for (const index of indices) {
			keep.add(index);
			let node = this.#nodes.get(index);
			if (!node) {
				node = this.#pool.pop() ?? (h(this.#doc, 'div', 'xg-hd') as HeaderNode);
				node.hidden = false;
				if (node.parentNode !== this.element) this.element.append(node);
				this.#nodes.set(index, node);
				node.xgSig = undefined;
			}
			const pos = row ? metrics.rowTop(index) : metrics.colLeft(index);
			const size = row ? metrics.rowHeight(index) : metrics.colWidth(index);
			const st = state(index);
			const sig = `${pos},${size},${st},${fontPx}`;
			if (node.xgSig === sig) continue;
			node.xgSig = sig;
			node.className = `xg-hd${st === 'part' ? ' xg-hd-sel' : st === 'full' ? ' xg-hd-full' : ''}`;
			node.style.cssText = row
				? `top:${pos}px;height:${size}px;font-size:${fontPx}px`
				: `left:${pos}px;width:${size}px;font-size:${fontPx}px`;
			node.textContent = row ? String(index + 1) : columnLabel(index);
			node.dataset.index = String(index);
			if (row) node.setAttribute('role', 'rowheader');
			else node.setAttribute('role', 'columnheader');
		}
		for (const [index, node] of this.#nodes) {
			if (keep.has(index)) continue;
			this.#nodes.delete(index);
			node.hidden = true;
			node.xgSig = undefined;
			this.#pool.push(node);
		}
	}
}
