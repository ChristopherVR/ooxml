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
	CHART_FILL_PARTS,
	type ChartFillPart,
	type ChartObject,
	type DrawingObject,
} from 'ooxml-core/xlsx';
import { h, place, svgNode, viewOf } from './dom';
import { QUADRANTS, type Box } from 'ooxml-core/xlsx/ui';
import type { GridView } from './grid-view';
import { paintTimeline, periodAt, timelineOf } from './timeline';
import { paintSlicer, slicerClickResult, slicerOf, type SlicerUiState } from './slicer';
import { paintSmartArt } from './smartart';
import { selectionOfPeriods } from 'ooxml-core/xlsx';
import { applyDragDelta, applyResize, type ResizeHandleId } from 'ooxml-core/geometry';
import { trackPointerGesture } from '../../pointer-gesture';

interface ObjectNode extends HTMLDivElement {
	xgSig?: string | undefined;
	xgSeriesHit?: number | undefined;
	xgChartPart?: ChartFillPart | undefined;
}

export class DrawingLayer {
	#view: GridView;
	#nodes = new Map<string, ObjectNode>();
	#urls = new Map<string, string>();
	#generation = 0;
	#cancelDrag: (() => void) | undefined;
	#slicerStates = new Map<string, SlicerUiState>();

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
		this.#cancelDrag?.();
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
					node.addEventListener('change', (event) => {
						const frame = view.sheet()?.drawings[index];
						const list = (event.target as Element).closest<HTMLSelectElement>(
							'[data-timeline-action=level]',
						);
						const session = view.ctx.session();
						if (list && frame?.kind === 'timeline' && session && !view.ctx.readOnly())
							this.#guard(() => session.updateTimeline(frame.name, { level: Number(list.value) }));
					});
					// Enter or Space on a focused slicer button arrives as a click without a pointer.
					node.addEventListener('click', (event) => {
						const frame = view.sheet()?.drawings[index];
						if (event.detail === 0 && frame?.kind === 'slicer')
							this.#pressSlicer(event, index, frame, false);
					});
					node.addEventListener('dblclick', (event) => {
						// Pointer capture retargets dblclick to the drawing container. Keep the
						// actual hit from pointerdown so a background click stays distinct.
						const frame = view.sheet()?.drawings[index];
						if (frame?.kind === 'smartArt') {
							// Excel opens the text pane on a double-click into a SmartArt graphic.
							event.stopPropagation();
							view.ctx.selection.set({ drawing: index });
							const textPane = view.ctx.commands.get('smartart.textPane');
							if (textPane && !textPane.checked?.(view.ctx))
								void view.ctx.commands.run('smartart.textPane');
							return;
						}
						const series = (event.currentTarget as ObjectNode).xgSeriesHit;
						const part = (event.currentTarget as ObjectNode).xgChartPart;
						if (series === undefined && part === undefined) return;
						event.stopPropagation();
						view.ctx.selection.set({ drawing: index });
						if (series !== undefined) void view.ctx.commands.run('chart.format-series', series);
						else void view.ctx.commands.run('chart.format-area', part);
					});
					this.#nodes.set(key, node);
				}
				if (node.parentNode !== parent) parent.append(node);
				place(node, box.x, box.y, box.w, box.h);
				const multi = drawing.kind === 'slicer' && this.#slicerState(drawing.name).multi;
				const sig = `${drawing.kind}:${Math.round(box.w)}x${Math.round(box.h)}:${this.#generation}:${multi}`;
				if (node.xgSig !== sig) {
					node.xgSig = sig;
					this.#paint(node, drawing, box);
				}
				this.#selection(node, index);
			}
		});
		for (const [key, node] of this.#nodes)
			if (!keep.has(key)) {
				node.remove();
				this.#nodes.delete(key);
			}
	}

	#paint(node: HTMLElement, drawing: DrawingObject, box: Box): void {
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
			this.#paintChart(node, drawing, box);
		} else if (drawing.kind === 'slicer') {
			paintSlicer(node, drawing, box.w, box.h, {
				ctx: view.ctx,
				workbook: view.workbook(),
				state: this.#slicerState(drawing.name),
			});
		} else if (drawing.kind === 'timeline') {
			paintTimeline(node, drawing, box.w, box.h, view.ctx, view.workbook());
		} else if (drawing.kind === 'smartArt') {
			paintSmartArt(view.ctx, node, drawing, view.workbook()?.theme, box);
		} else if (drawing.chartEx) {
			// An extended chart (waterfall, histogram) the core painter draws; the frame is read-only.
			this.#paintChart(node, drawing.chartEx, box);
		} else {
			node.classList.add('xg-obj-unsupported');
			node.textContent = view.ctx.t('{name} (not shown)', { name: drawing.description });
		}
	}

	#paintChart(node: HTMLElement, drawing: ChartObject, box: Box): void {
		const view = this.#view;
		const doc = view.doc;
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
						renderChartSvg(model, Math.max(10, box.w), Math.max(10, box.h), {
							measureText: (content, font) => view.measurer.measure(content, font),
							measureFont: (font) => view.measurer.fontMetrics(font),
						}),
					),
				);
			} catch {
				node.textContent = view.ctx.t('This chart could not be displayed.');
			}
		}
		node.setAttribute('role', 'img');
		node.setAttribute('aria-label', drawing.title ?? view.ctx.t('Chart'));
	}

	#selection(node: HTMLElement, index: number): void {
		const selected = this.#selected === index;
		node.classList.toggle('xg-obj-sel', selected);
		if (!selected) {
			for (const grip of node.querySelectorAll('.xg-grip')) grip.remove();
		} else if (!node.querySelector('.xg-grip')) {
			for (const pos of ['nw', 'ne', 'sw', 'se']) {
				const grip = h(this.#view.doc, 'div', 'xg-grip', { 'data-grip': pos });
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
		this.#cancelDrag?.();
		// The timeline's level list must keep its native pointer behaviour (it opens on mousedown).
		if (!(event.target as Element).closest?.('select')) event.preventDefault();
		const sheet = view.sheet();
		const drawing = sheet?.drawings[index];
		if (!sheet || !drawing) return;
		const node = event.currentTarget as ObjectNode;
		if (drawing.kind === 'slicer' && this.#pressSlicer(event, index, drawing)) return;
		if (drawing.kind === 'timeline' && this.#pressTimeline(event, index, drawing, node)) return;
		const series = (event.target as Element).closest<SVGElement>('[data-chart-series]')?.dataset
			.chartSeries;
		node.xgSeriesHit = series === undefined ? undefined : Number(series);
		const part = (event.target as Element).closest<SVGElement>('[data-chart-part]')?.dataset
			.chartPart;
		node.xgChartPart = CHART_FILL_PARTS.includes(part as ChartFillPart)
			? (part as ChartFillPart)
			: undefined;
		view.ctx.selection.set({ drawing: index });
		view.ctx.grid()?.focus();
		const session = view.ctx.session();
		if (view.ctx.readOnly() || !session) return;
		const grip = (event.target as Element).closest<HTMLElement>('[data-grip]')?.dataset.grip as
			| ResizeHandleId
			| undefined;
		const start = view.clientToView(event.clientX, event.clientY);
		const origin = anchorToPixelBox(sheet, drawing.anchor, view.metrics);
		let box = origin;
		const move = (e: PointerEvent) => {
			const p = view.clientToView(e.clientX, e.clientY);
			const dx = p.x - start.x;
			const dy = p.y - start.y;
			const startBox = { x: origin.x, y: origin.y, width: origin.w, height: origin.h };
			const next = grip
				? applyResize(startBox, grip, dx, dy, 1, { minSize: 8 })
				: applyDragDelta(startBox, dx, dy, 1);
			box = { x: Math.max(0, next.x), y: Math.max(0, next.y), w: next.width, h: next.height };
			place(node, box.x, box.y, box.w, box.h);
		};
		const up = () => {
			this.#cancelDrag = undefined;
			if (box.x !== origin.x || box.y !== origin.y || box.w !== origin.w || box.h !== origin.h)
				session.setDrawingAnchor(
					view.sheetIndex(),
					index,
					pixelBoxToAnchor(sheet, box, view.metrics, drawing.anchor),
				);
		};
		this.#cancelDrag = trackPointerGesture(node, event.pointerId, {
			move,
			commit: up,
			cancel: () => {
				this.#cancelDrag = undefined;
				place(node, origin.x, origin.y, origin.w, origin.h);
			},
		});
	}

	/** Runs an edit from a timeline or slicer control and reports a refusal as a toast. */
	#guard(run: () => void): void {
		try {
			run();
		} catch (error) {
			this.#view.ctx.toast(
				this.#view.ctx.t(error instanceof Error ? error.message : String(error)),
				'warning',
			);
		}
	}

	/** Handles a press on a timeline control; false lets the press move or resize the frame. */
	#pressTimeline(
		event: PointerEvent,
		index: number,
		frame: Extract<DrawingObject, { kind: 'timeline' }>,
		node: ObjectNode,
	): boolean {
		const target = event.target as Element;
		const view = this.#view;
		const list = target.closest('[data-timeline-action=level]');
		const clear = target.closest('[data-timeline-action=clear]');
		const cell = target.closest<HTMLElement>('[data-period]');
		if (!list && !clear && !cell) return false;
		view.ctx.selection.set({ drawing: index });
		if (list) return true;
		view.ctx.grid()?.focus();
		const found = timelineOf(view.workbook(), frame, view.ctx.locale());
		const session = view.ctx.session();
		if (!found || !session || view.ctx.readOnly()) return true;
		if (clear) {
			this.#guard(() => session.setTimelineSelection(found.cache.name, null));
			return true;
		}
		const track = cell!.closest<HTMLElement>('.xg-tl-track');
		if (!track) return true;
		const last = found.view.periods.length - 1;
		const start = Number(cell!.dataset.period);
		const anchor = event.shiftKey && found.view.selected ? found.view.selected[0] : start;
		let current = start;
		const cells = [...track.querySelectorAll<HTMLElement>('[data-period]')];
		const preview = () => {
			const [from, to] = [Math.min(anchor, current), Math.max(anchor, current)];
			for (const el of cells) {
				const i = Number(el.dataset.period);
				el.classList.toggle('xg-tl-on', i >= from && i <= to);
			}
		};
		preview();
		node.setPointerCapture?.(event.pointerId);
		trackPointerGesture(node, event.pointerId, {
			move: (move) => {
				current = Math.max(0, Math.min(last, periodAt(track, move.clientX)));
				preview();
			},
			commit: () => {
				const only = found.view.selected;
				const same = only && only[0] === start && only[1] === start && anchor === current;
				const selection = same ? null : selectionOfPeriods(found.view.periods, anchor, current);
				this.#guard(() => session.setTimelineSelection(found.cache.name, selection ?? null));
				this.invalidate();
				this.render();
			},
			cancel: () => {
				this.invalidate();
				this.render();
			},
		});
		return true;
	}

	#slicerState(name: string): SlicerUiState {
		let state = this.#slicerStates.get(name);
		if (!state) this.#slicerStates.set(name, (state = { multi: false }));
		return state;
	}

	/** Handles a press on a slicer button; false lets the press move or resize the slicer. */
	#pressSlicer(
		event: MouseEvent,
		index: number,
		frame: Extract<DrawingObject, { kind: 'slicer' }>,
		pointer = true,
	): boolean {
		const target = event.target as Element;
		const view = this.#view;
		const item = target.closest<HTMLElement>('[data-slicer-item]');
		const tool = target.closest<HTMLElement>('[data-slicer-action]');
		if (!item && !tool) return false;
		view.ctx.selection.set({ drawing: index });
		if (pointer) view.ctx.grid()?.focus();
		const found = slicerOf(view.workbook(), frame);
		const session = view.ctx.session();
		const state = this.#slicerState(frame.name);
		if (!found || !session || view.ctx.readOnly()) return true;
		try {
			if (tool?.dataset.slicerAction === 'multi') {
				state.multi = !state.multi;
				this.invalidate();
				this.render();
			} else if (tool?.dataset.slicerAction === 'clear') {
				delete state.anchor;
				session.setSlicerSelection(found.cache.name, null);
			} else if (item?.dataset.slicerItem !== undefined) {
				const key = item.dataset.slicerItem;
				const keys = slicerClickResult(found.items, key, event, state);
				state.anchor = key;
				session.setSlicerSelection(found.cache.name, keys);
			}
		} catch (error) {
			view.ctx.toast(view.ctx.t(error instanceof Error ? error.message : String(error)), 'warning');
		}
		return true;
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
		this.#cancelDrag?.();
		const win = viewOf(this.#view.root);
		for (const url of this.#urls.values()) win?.URL?.revokeObjectURL?.(url);
		this.#urls.clear();
		for (const node of this.#nodes.values()) node.remove();
		this.#nodes.clear();
	}
}
