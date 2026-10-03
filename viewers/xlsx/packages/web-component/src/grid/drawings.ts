// Pictures, charts and SmartArt over the grid, positioned from their anchors in the quadrant of their
// top-left cell. Pictures come from the package parts as object URLs; charts are drawn by the core
// (`chartView` + `renderChartSvg`) from live cell values. When editing, objects can be selected,
// moved and resized (the anchor change is one undoable edit). The selected object lives in the
// shared selection model (`selection.drawing`), so the Chart Design tab and Delete follow it.
import {
	anchorToPixelBox,
	chartView,
	createRefEvaluator,
	pixelBoxToAnchor,
	renderChartSvg,
	type DrawingObject,
} from '@christophervr/xlsx-core';
import { h, place, svgNode, viewOf } from './dom.js';
import { QUADRANTS, type Box } from './geometry.js';
import type { GridView } from './grid-view.js';
import { paintSmartArt } from './smartart.js';

interface ObjectNode extends HTMLDivElement {
	xgSig?: string | undefined;
}

export class DrawingLayer {
	#view: GridView;
	#nodes = new Map<string, ObjectNode>();
	#urls = new Map<string, string>();
	#generation = 0;

	constructor(view: GridView) {
		this.#view = view;
		view.hooks.add(() => this.render());
	}

	/** The selected drawing of the shown sheet, from the selection model. */
	get #selected(): number | undefined {
		const selection = this.#view.ctx.selection.get();
		return selection.sheet === this.#view.sheetIndex() ? selection.drawing : undefined;
	}

	/** Charts depend on cell values: bump after model changes so they redraw. */
	invalidate(): void {
		this.#generation++;
	}

	#url(partName: string, contentType: string): string | undefined {
		const cached = this.#urls.get(partName);
		if (cached) return cached;
		const bytes = this.#view.workbook()?.source?.parts.get(partName);
		const win = viewOf(this.#view.root);
		if (!bytes || !win?.URL?.createObjectURL) return undefined;
		const url = win.URL.createObjectURL(new Blob([bytes.slice().buffer], { type: contentType }));
		this.#urls.set(partName, url);
		return url;
	}

	render(): void {
		const view = this.#view;
		const sheet = view.sheet();
		const keep = new Set<string>();
		const g = view.geometry;
		sheet?.drawings.forEach((drawing, index) => {
			const box = anchorToPixelBox(sheet, drawing.anchor, view.metrics);
			// An object may cross the frozen-pane split: paint it in every pane (each clips its part).
			for (const quadrant of QUADRANTS) {
				const pane = g.box(quadrant);
				if (pane.w <= 0 || pane.h <= 0) continue;
				const key = `${view.sheetIndex()}:${index}:${quadrant}`;
				keep.add(key);
				const parent = view.quads[quadrant].drawings;
				let node = this.#nodes.get(key);
				if (!node) {
					node = h(view.doc, 'div', 'xg-obj') as ObjectNode;
					node.dataset.index = String(index);
					node.addEventListener('pointerdown', (event) => this.#press(event, index));
					this.#nodes.set(key, node);
				}
				if (node.parentNode !== parent) parent.append(node);
				place(node, box.x, box.y, box.w, box.h);
				const sig = `${drawing.kind}:${Math.round(box.w)}x${Math.round(box.h)}:${this.#generation}:${this.#selected === index}`;
				if (node.xgSig === sig) continue;
				node.xgSig = sig;
				this.#paint(node, drawing, box, index);
			}
		});
		for (const [key, node] of this.#nodes)
			if (!keep.has(key)) {
				node.remove();
				this.#nodes.delete(key);
			}
	}

	#paint(node: HTMLElement, drawing: DrawingObject, box: Box, index: number): void {
		const view = this.#view;
		const doc = view.doc;
		node.replaceChildren();
		node.className = 'xg-obj';
		if (drawing.kind === 'image') {
			const img = h(doc, 'img', '', {
				alt: drawing.description ?? drawing.name ?? '',
				draggable: 'false',
			});
			const url = this.#url(drawing.partName, drawing.contentType);
			if (url) img.src = url;
			node.append(img);
			node.setAttribute('role', 'img');
		} else if (drawing.kind === 'chart') {
			node.classList.add('xg-chart');
			const workbook = view.workbook();
			if (workbook) {
				try {
					const evaluate = createRefEvaluator(workbook, view.ctx.session()?.calc, {
						sheet: view.sheetIndex(),
					});
					const model = chartView(workbook, view.sheetIndex(), drawing, evaluate);
					node.append(
						svgNode(
							doc,
							renderChartSvg(
								model,
								Math.max(10, Math.round(box.w)),
								Math.max(10, Math.round(box.h)),
							),
						),
					);
				} catch {
					node.textContent = view.ctx.t('This chart could not be displayed.');
				}
			}
			node.setAttribute('role', 'img');
			node.setAttribute('aria-label', drawing.title ?? view.ctx.t('Chart'));
		} else if (drawing.kind === 'smartArt') {
			paintSmartArt(view.ctx, node, drawing, view.workbook()?.theme);
		} else {
			node.classList.add('xg-obj-unsupported');
			node.textContent = view.ctx.t('{name} (not shown)', { name: drawing.description });
		}
		if (this.#selected === index) {
			node.classList.add('xg-obj-sel');
			for (const pos of ['nw', 'ne', 'sw', 'se']) {
				const grip = h(doc, 'div', 'xg-grip', { 'data-grip': pos });
				grip.style.left = pos.includes('w') ? '-4px' : 'calc(100% - 4px)';
				grip.style.top = pos.includes('n') ? '-4px' : 'calc(100% - 4px)';
				grip.style.cursor = pos === 'nw' || pos === 'se' ? 'nwse-resize' : 'nesw-resize';
				node.append(grip);
			}
		}
	}

	deselect(): void {
		if (this.#selected === undefined) return;
		this.#view.ctx.selection.set({ drawing: undefined });
	}

	#press(event: PointerEvent, index: number): void {
		const view = this.#view;
		event.stopPropagation();
		if (event.button !== 0) return;
		event.preventDefault();
		const sheet = view.sheet();
		const drawing = sheet?.drawings[index];
		if (!sheet || !drawing) return;
		view.ctx.selection.set({ drawing: index });
		view.ctx.grid()?.focus();
		const session = view.ctx.session();
		if (view.ctx.readOnly() || !session) return;
		const grip = (event.target as Element).closest<HTMLElement>('[data-grip]')?.dataset.grip;
		const start = view.clientToView(event.clientX, event.clientY);
		const origin = anchorToPixelBox(sheet, drawing.anchor, view.metrics);
		const node = event.currentTarget as HTMLElement;
		let box = origin;
		node.setPointerCapture?.(event.pointerId);
		const move = (e: PointerEvent) => {
			const p = view.clientToView(e.clientX, e.clientY);
			const dx = p.x - start.x;
			const dy = p.y - start.y;
			if (!grip) box = { ...origin, x: Math.max(0, origin.x + dx), y: Math.max(0, origin.y + dy) };
			else {
				const left = grip.includes('w')
					? Math.min(origin.x + dx, origin.x + origin.w - 8)
					: origin.x;
				const top = grip.includes('n')
					? Math.min(origin.y + dy, origin.y + origin.h - 8)
					: origin.y;
				const right = grip.includes('e')
					? Math.max(origin.x + origin.w + dx, left + 8)
					: origin.x + origin.w;
				const bottom = grip.includes('s')
					? Math.max(origin.y + origin.h + dy, top + 8)
					: origin.y + origin.h;
				box = { x: Math.max(0, left), y: Math.max(0, top), w: right - left, h: bottom - top };
			}
			place(node, box.x, box.y, box.w, box.h);
		};
		const up = (e: PointerEvent) => {
			node.removeEventListener('pointermove', move);
			node.removeEventListener('pointerup', up);
			node.releasePointerCapture?.(e.pointerId);
			if (box !== origin)
				session.setDrawingAnchor(
					view.sheetIndex(),
					index,
					pixelBoxToAnchor(sheet, box, view.metrics, drawing.anchor),
				);
		};
		node.addEventListener('pointermove', move);
		node.addEventListener('pointerup', up);
	}

	/** Moves the selected drawing by a few pixels (arrow keys); false when nothing is selected. */
	nudge(dx: number, dy: number): boolean {
		const view = this.#view;
		const index = this.#selected;
		const sheet = view.sheet();
		const drawing = index === undefined ? undefined : sheet?.drawings[index];
		const session = view.ctx.session();
		if (index === undefined || !sheet || !drawing) return false;
		if (!session || view.ctx.readOnly()) return true;
		const box = anchorToPixelBox(sheet, drawing.anchor, view.metrics);
		const next = { ...box, x: Math.max(0, box.x + dx), y: Math.max(0, box.y + dy) };
		session.setDrawingAnchor(
			view.sheetIndex(),
			index,
			pixelBoxToAnchor(sheet, next, view.metrics, drawing.anchor),
		);
		return true;
	}

	destroy(): void {
		const win = viewOf(this.#view.root);
		for (const url of this.#urls.values()) win?.URL?.revokeObjectURL?.(url);
		this.#urls.clear();
		for (const node of this.#nodes.values()) node.remove();
		this.#nodes.clear();
	}
}
