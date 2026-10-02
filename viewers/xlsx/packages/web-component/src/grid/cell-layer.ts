// DOM painter for one quadrant's cells. Nodes are keyed by cell and recycled through a pool, and
// a node is only rewritten when its paint signature changed, so scrolling touches the rows and
// columns that came into view and an edit touches the cells whose view changed.
import type { CellView } from '@christophervr/xlsx-core';
import {
	alignItems,
	cssFont,
	diagonalSvg,
	edgeCss,
	fillPaint,
	hasBorders,
	justify,
	textAlign,
	textDecoration,
} from './cell-paint.js';
import type { CellItem } from './cell-items.js';
import { h, svgNode } from './dom.js';
import { iconSvg } from './icon-sets.js';

interface CellNode extends HTMLDivElement {
	xgSig?: string | undefined;
}

const viewKeys = new WeakMap<CellView, string>();
const viewKey = (view: CellView): string => {
	let key = viewKeys.get(view);
	if (key === undefined) {
		key = JSON.stringify(view);
		viewKeys.set(view, key);
	}
	return key;
};

export function itemSignature(item: CellItem, zoom: number): string {
	return `${item.x},${item.y},${item.w},${item.h},${item.textLeft},${item.textWidth},${item.fontScale},${item.errorMark ? 1 : 0},${zoom}\u0000${item.text}\u0000${viewKey(item.view)}`;
}

export class CellLayer {
	readonly element: HTMLDivElement;
	readonly #doc: Document;
	readonly #nodes = new Map<string, CellNode>();
	readonly #pool: CellNode[] = [];

	constructor(doc: Document) {
		this.#doc = doc;
		this.element = h(doc, 'div', 'xg-cells');
	}

	get size(): number {
		return this.#nodes.size;
	}

	/** Paints `items`; nodes not in the list go back to the pool. */
	render(items: readonly CellItem[], zoom: number): void {
		const keep = new Set<string>();
		for (const item of items) {
			keep.add(item.key);
			let node = this.#nodes.get(item.key);
			if (!node) {
				node = this.#pool.pop() ?? (h(this.#doc, 'div', 'xg-c') as CellNode);
				node.xgSig = undefined;
				this.#nodes.set(item.key, node);
				if (node.parentNode !== this.element) this.element.append(node);
				node.hidden = false;
			}
			const sig = itemSignature(item, zoom);
			if (node.xgSig !== sig) {
				paintCell(this.#doc, node, item, zoom);
				node.xgSig = sig;
			}
		}
		for (const [key, node] of this.#nodes) {
			if (keep.has(key)) continue;
			this.#nodes.delete(key);
			node.hidden = true;
			node.xgSig = undefined;
			if (this.#pool.length < 400) this.#pool.push(node);
			else node.remove();
		}
	}

	nodeAt(row: number, col: number): HTMLElement | undefined {
		return this.#nodes.get(`${row}:${col}`);
	}

	clear(): void {
		this.render([], 100);
	}
}

function paintText(doc: Document, box: HTMLElement, item: CellItem, zoom: number): void {
	const { view } = item;
	const font = view.font;
	const style = box.style;
	style.left = `${item.textLeft}px`;
	style.width = `${item.textWidth}px`;
	style.font = cssFont(font, zoom * item.fontScale);
	style.color = font.color;
	style.textDecoration = textDecoration(font);
	style.justifyContent = justify(view.hAlign);
	style.alignItems = alignItems(view.vAlign);
	style.textAlign = textAlign(view.hAlign);
	style.whiteSpace = view.wrap ? 'pre-wrap' : 'pre';
	const pad = Math.max(1, Math.round(2 * (zoom / 100)));
	const indent = pad + view.indentPx * (zoom / 100);
	style.paddingLeft = `${view.hAlign === 'left' || view.hAlign === 'distributed' ? indent : pad}px`;
	style.paddingRight = `${view.hAlign === 'right' ? indent : pad}px`;
	box.classList.toggle('xg-wrap', view.wrap);
	box.classList.toggle('xg-vert', Boolean(view.verticalText));
	box.classList.toggle('xg-over', item.overflowing);
	const span = h(doc, 'span', 'xg-tx');
	if (view.rotation) span.style.transform = `rotate(${-view.rotation}deg)`;
	if (font.vertAlign) span.style.verticalAlign = font.vertAlign === 'super' ? 'super' : 'sub';
	if (view.rich?.length && item.text === view.text) {
		for (const run of view.rich) {
			const part = h(doc, 'span');
			part.textContent = run.text;
			part.style.font = cssFont(run.font, zoom * item.fontScale);
			part.style.color = run.font.color;
			part.style.textDecoration = textDecoration(run.font);
			if (run.font.vertAlign)
				part.style.verticalAlign = run.font.vertAlign === 'super' ? 'super' : 'sub';
			span.append(part);
		}
	} else span.textContent = item.text;
	box.replaceChildren(span);
}

function paintBorders(doc: Document, node: HTMLElement, item: CellItem): void {
	const b = item.view.borders;
	if (!hasBorders(b)) return;
	const box = h(doc, 'div', 'xg-b');
	box.style.width = `${item.w + 1}px`;
	box.style.height = `${item.h + 1}px`;
	if (b.top) box.style.borderTop = edgeCss(b.top);
	if (b.right) box.style.borderRight = edgeCss(b.right);
	if (b.bottom) box.style.borderBottom = edgeCss(b.bottom);
	if (b.left) box.style.borderLeft = edgeCss(b.left);
	node.append(box);
	const diagonal = diagonalSvg(b, item.w, item.h);
	if (diagonal) {
		const svg = svgNode(doc, diagonal);
		svg.setAttribute('class', 'xg-diag');
		node.append(svg);
	}
}

/** Writes a cell node from scratch (only called when its signature changed). */
export function paintCell(doc: Document, node: HTMLElement, item: CellItem, zoom: number): void {
	const { view } = item;
	const style = node.style;
	style.left = `${item.x}px`;
	style.top = `${item.y}px`;
	style.width = `${item.w}px`;
	style.height = `${item.h}px`;
	const fill = fillPaint(view.fill);
	style.backgroundColor = fill?.background ?? (item.merged ? 'var(--xve-sheet-bg, #fff)' : '');
	style.backgroundImage = fill?.image ?? '';
	node.className = `xg-c${item.overflowing ? ' xg-c-over' : ''}${view.hasHyperlink ? ' xg-link' : ''}${item.merged ? ' xg-merged' : ''}`;
	node.replaceChildren();
	if (view.dataBar) {
		const bar = h(doc, 'div', 'xg-db');
		const width = Math.max(0, Math.min(1, view.dataBar.fraction)) * (item.w - 4);
		bar.style.width = `${width}px`;
		const color = view.dataBar.negative ? '#ff0000' : view.dataBar.color;
		bar.style.background = `linear-gradient(90deg, ${color}, color-mix(in srgb, ${color} 15%, #fff))`;
		bar.style.borderColor = color;
		node.append(bar);
	}
	if (view.icon) {
		const icon = svgNode(
			doc,
			iconSvg(view.icon.set, view.icon.index, Math.round(16 * (zoom / 100))),
		);
		icon.setAttribute('class', 'xg-ic');
		node.append(icon);
	}
	if (item.text) {
		const box = h(doc, 'div', 'xg-t');
		paintText(doc, box, item, zoom);
		if (view.icon && view.hAlign !== 'right') box.style.paddingLeft = `${20 * (zoom / 100)}px`;
		node.append(box);
	}
	paintBorders(doc, node, item);
	if (view.hasComment) node.append(h(doc, 'div', 'xg-cm'));
	if (item.errorMark) node.append(h(doc, 'div', 'xg-er'));
}
