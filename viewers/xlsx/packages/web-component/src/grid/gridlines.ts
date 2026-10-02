// Gridlines of one quadrant: one thin line per visible row and column (a few dozen nodes),
// reused by index. Filled and merged cells paint over them, as in Excel.
import type { GridMetrics } from '@christophervr/xlsx-core';
import { h } from './dom.js';

export class GridlineLayer {
	readonly element: HTMLDivElement;
	readonly #doc: Document;
	readonly #lines: HTMLDivElement[] = [];

	constructor(doc: Document) {
		this.#doc = doc;
		this.element = h(doc, 'div', 'xg-lines');
	}

	render(
		metrics: GridMetrics,
		rows: readonly number[],
		cols: readonly number[],
		visible: boolean,
	): void {
		this.element.hidden = !visible;
		if (!visible || !rows.length || !cols.length) {
			for (const line of this.#lines) line.hidden = true;
			return;
		}
		const first = rows[0] ?? 0;
		const last = rows[rows.length - 1] ?? first;
		const left = metrics.colLeft(cols[0] ?? 0);
		const right = metrics.colLeft((cols[cols.length - 1] ?? 0) + 1);
		const top = metrics.rowTop(first);
		const bottom = metrics.rowTop(last + 1);
		let i = 0;
		const line = (x: number, y: number, w: number, hgt: number, vertical: boolean) => {
			let node = this.#lines[i];
			if (!node) {
				node = h(this.#doc, 'div');
				this.#lines.push(node);
				this.element.append(node);
			}
			node.className = vertical ? 'xg-gl xg-gl-v' : 'xg-gl xg-gl-h';
			node.hidden = false;
			node.style.cssText = `left:${x}px;top:${y}px;width:${w}px;height:${hgt}px`;
			i++;
		};
		for (const col of cols) line(metrics.colLeft(col + 1) - 1, top, 1, bottom - top, true);
		for (const row of rows) line(left, metrics.rowTop(row + 1) - 1, right - left, 1, false);
		for (let j = i; j < this.#lines.length; j++) {
			const node = this.#lines[j];
			if (node) node.hidden = true;
		}
	}
}
