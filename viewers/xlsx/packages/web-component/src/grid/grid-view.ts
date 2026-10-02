// The virtualized grid: a native scroller whose sticky view holds four quadrant planes (frozen
// corner, frozen rows, frozen columns, scrolling cells) plus row/column headers. Only visible
// rows and columns get nodes; a scroll moves plane transforms and paints the cells that came
// into view (one paint per animation frame).
import {
	createWorksheet,
	createGridMetrics,
	getCell,
	mergeAt,
	visibleCells,
	type CellRange,
	type CellView,
	type GridMetrics,
	type Workbook,
	type Worksheet,
} from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { buildItems } from './cell-items.js';
import { fitNumber } from './number-fit.js';
import { CellLayer } from './cell-layer.js';
import { h, nextFrame, place } from './dom.js';
import {
	columnHeaderHeight,
	GridGeometry,
	QUADRANTS,
	rowHeaderWidth,
	scrollToReveal,
	type Hit,
	type Quadrant,
} from './geometry.js';
import { GridlineLayer } from './gridlines.js';
import { paintChrome, type HeaderPane } from './pane-chrome.js';
import { CellViewCache } from './view-cache.js';
import { HeaderLayer } from './headers.js';
import { createTextMeasurer, type TextMeasurer } from './measure.js';
import { SelectionLayer, type OverlayState } from './selection-layer.js';
import { cellRange } from './selection-ops.js';

interface QuadrantNodes {
	box: HTMLDivElement;
	plane: HTMLDivElement;
	cells: CellLayer;
	lines: GridlineLayer;
	overlay: SelectionLayer;
	drawings: HTMLDivElement;
}

const EMPTY_SHEET = createWorksheet('Sheet1', 1);

export type PaintHook = (view: GridView) => void;

export class GridView {
	readonly ctx: EditorContext;
	readonly doc: Document;
	readonly root: HTMLDivElement;
	readonly scroller: HTMLDivElement;
	readonly space: HTMLDivElement;
	readonly view: HTMLDivElement;
	readonly corner: HTMLDivElement;
	readonly quads = {} as Record<Quadrant, QuadrantNodes>;
	readonly colHeaders: { frozen: HeaderPane; scroll: HeaderPane };
	readonly rowHeaders: { frozen: HeaderPane; scroll: HeaderPane };
	readonly measurer: TextMeasurer;
	readonly overlay: Pick<OverlayState, 'copyRange' | 'references' | 'fillPreview'> = {};
	readonly hooks = new Set<PaintHook>();
	metrics!: GridMetrics;
	geometry!: GridGeometry;
	extentRow = 0;
	extentCol = 0;
	#cache: CellViewCache = new CellViewCache({
		workbook: () => this.workbook(),
		sheet: () => this.sheet(),
		sheetIndex: () => this.sheetIndex(),
		session: () => this.ctx.session(),
	});
	#cancelFrame: (() => void) | undefined;
	#zoomOverride = new Map<Worksheet, number>();
	#freezeH: HTMLDivElement;
	#freezeV: HTMLDivElement;
	#sheetRef: Worksheet | undefined;

	constructor(ctx: EditorContext, container: HTMLElement) {
		this.ctx = ctx;
		this.doc = container.ownerDocument;
		const doc = this.doc;
		this.measurer = createTextMeasurer(doc);
		this.root = h(doc, 'div', 'xg-root', { role: 'grid', part: 'grid' });
		this.scroller = h(doc, 'div', 'xg-scroller');
		this.space = h(doc, 'div', 'xg-space');
		this.view = h(doc, 'div', 'xg-view');
		this.space.append(this.view);
		this.scroller.append(this.space);
		this.root.append(this.scroller);
		for (const q of QUADRANTS) {
			const box = h(doc, 'div', `xg-q xg-q-${q}`);
			const plane = h(doc, 'div', 'xg-plane');
			const cells = new CellLayer(doc);
			const lines = new GridlineLayer(doc);
			const overlay = new SelectionLayer(doc);
			const drawings = h(doc, 'div', 'xg-drawings');
			plane.append(lines.element, cells.element, overlay.element, drawings);
			box.append(plane);
			this.view.append(box);
			this.quads[q] = { box, plane, cells, lines, overlay, drawings };
		}
		const pane = (axis: 'row' | 'col'): HeaderPane => {
			const box = h(doc, 'div', `xg-hbox xg-hbox-${axis}`);
			const layer = new HeaderLayer(doc, axis);
			box.append(layer.element);
			this.view.append(box);
			return { box, layer };
		};
		this.colHeaders = { frozen: pane('col'), scroll: pane('col') };
		this.rowHeaders = { frozen: pane('row'), scroll: pane('row') };
		this.corner = h(doc, 'div', 'xg-corner', { role: 'button', title: ctx.t('Select All') });
		this.#freezeH = h(doc, 'div', 'xg-freeze-h');
		this.#freezeV = h(doc, 'div', 'xg-freeze-v');
		this.view.append(this.corner, this.#freezeH, this.#freezeV);
		container.append(this.root);
		this.scroller.addEventListener('scroll', () => this.schedule(), { passive: true });
		this.rebuild();
	}

	workbook(): Workbook | undefined {
		return this.ctx.workbook();
	}
	sheetIndex(): number {
		return this.ctx.activeSheet();
	}
	sheet(): Worksheet | undefined {
		return this.workbook()?.sheets[this.sheetIndex()];
	}

	zoom(): number {
		const sheet = this.sheet();
		return (sheet && this.#zoomOverride.get(sheet)) ?? sheet?.view.zoom ?? 100;
	}
	setZoom(percent: number): void {
		const sheet = this.sheet();
		if (!sheet) return;
		this.#zoomOverride.set(sheet, Math.max(10, Math.min(400, Math.round(percent))));
		this.rebuild();
	}

	/** Recreates metrics (structure, sizes, zoom or sheet changed) and repaints. */
	rebuild(): void {
		const sheet = this.sheet();
		const same = sheet === this.#sheetRef;
		if (!same) {
			this.#sheetRef = sheet;
			for (const q of QUADRANTS) this.quads[q].cells.clear();
		}
		this.metrics = createGridMetrics(sheet ?? EMPTY_SHEET, { zoom: this.zoom() });
		this.extentRow = same ? Math.max(this.metrics.lastRow, this.extentRow) : this.metrics.lastRow;
		this.extentCol = same ? Math.max(this.metrics.lastCol, this.extentCol) : this.metrics.lastCol;
		this.clearViews();
		this.#layout();
		this.paint();
	}

	/** Drops cached cell views (after any model change) and schedules a repaint. */
	invalidate(): void {
		this.clearViews();
		this.schedule();
	}

	clearViews(): void {
		this.#cache.clear();
	}

	/** Grows the scroll extent so `row`/`col` (and some margin) are reachable. */
	ensureExtent(row: number, col: number): void {
		const r = Math.max(this.extentRow, row + 30);
		const c = Math.max(this.extentCol, col + 10);
		if (r === this.extentRow && c === this.extentCol) return;
		this.extentRow = Math.min(r, 1_048_575);
		this.extentCol = Math.min(c, 16_383);
		this.#layout();
	}

	headerSizes(): { w: number; h: number } {
		if (this.sheet()?.view.showHeaders === false) return { w: 0, h: 0 };
		const zoom = this.metrics.zoom;
		return { w: rowHeaderWidth(this.extentRow, zoom), h: columnHeaderHeight(zoom) };
	}

	#layout(): void {
		const { w, h: hh } = this.headerSizes();
		this.space.style.width = `${w + this.metrics.totalWidth(this.extentCol)}px`;
		this.space.style.height = `${hh + this.metrics.totalHeight(this.extentRow)}px`;
	}

	#geometry(): GridGeometry {
		const { w, h: hh } = this.headerSizes();
		const width = this.scroller.clientWidth || this.root.clientWidth || 800;
		const height = this.scroller.clientHeight || this.root.clientHeight || 600;
		return new GridGeometry({
			metrics: this.metrics,
			freeze: this.sheet()?.view.freeze,
			headerW: w,
			headerH: hh,
			width,
			height,
			scrollLeft: this.scroller.scrollLeft,
			scrollTop: this.scroller.scrollTop,
		});
	}

	schedule(): void {
		if (this.#cancelFrame) return;
		this.#cancelFrame = nextFrame(this.root, () => {
			this.#cancelFrame = undefined;
			this.paint();
		});
	}

	/** The cached core view of a cell; null when it paints nothing. */
	cellViewAt(row: number, col: number): CellView | null {
		return this.#cache.get(row, col);
	}

	/** The active cell's range, merge-expanded. */
	activeRange(): CellRange {
		const active = this.ctx.selection.get().active;
		const sheet = this.sheet();
		return (sheet && mergeAt(sheet, active.row, active.col)) ?? cellRange(active);
	}

	clientToView(clientX: number, clientY: number): { x: number; y: number } {
		const rect = this.view.getBoundingClientRect();
		return { x: clientX - rect.left, y: clientY - rect.top };
	}

	hitClient(clientX: number, clientY: number): Hit {
		const { x, y } = this.clientToView(clientX, clientY);
		return this.geometry.hit(x, y);
	}

	/** Scrolls so a cell is fully visible in the scrolling pane. */
	reveal(row: number, col: number): void {
		this.ensureExtent(row, col);
		this.geometry = this.#geometry();
		const next = scrollToReveal(this.geometry, row, col);
		if (next.left !== this.scroller.scrollLeft) this.scroller.scrollLeft = next.left;
		if (next.top !== this.scroller.scrollTop) this.scroller.scrollTop = next.top;
		this.schedule();
	}

	paint(): void {
		this.#cancelFrame?.();
		this.#cancelFrame = undefined;
		const g = (this.geometry = this.#geometry());
		const sheet = this.sheet();
		this.view.style.width = `${g.width}px`;
		this.view.style.height = `${g.height}px`;
		if (g.scrollTop + g.height > this.metrics.totalHeight(this.extentRow) - 400)
			this.ensureExtent(this.extentRow + 200, 0);
		if (g.scrollLeft + g.width > this.metrics.totalWidth(this.extentCol) - 300)
			this.ensureExtent(0, this.extentCol + 10);
		const freeze = sheet?.view.freeze;
		const vis = visibleCells(
			this.metrics,
			{
				scrollLeft: g.scrollLeft,
				scrollTop: g.scrollTop,
				width: g.cellsW + 40,
				height: g.cellsH + 30,
			},
			freeze,
		);
		const zoom = this.metrics.zoom;
		const selection = this.ctx.selection.get();
		const editable = !this.ctx.readOnly();
		const gridlines = sheet?.view.showGridLines !== false;
		const sets: Record<Quadrant, [readonly number[], readonly number[]]> = {
			main: [vis.rows, vis.cols],
			top: [vis.frozenRows, vis.cols],
			left: [vis.rows, vis.frozenCols],
			corner: [vis.frozenRows, vis.frozenCols],
		};
		const maxRow =
			Math.max(vis.rows[vis.rows.length - 1] ?? 0, vis.frozenRows[vis.frozenRows.length - 1] ?? 0) +
			1;
		const maxCol =
			Math.max(vis.cols[vis.cols.length - 1] ?? 0, vis.frozenCols[vis.frozenCols.length - 1] ?? 0) +
			1;
		const activeRange = this.activeRange();
		for (const q of QUADRANTS) {
			const nodes = this.quads[q];
			const box = g.box(q);
			const [rows, cols] = sets[q];
			nodes.box.hidden = box.w <= 0 || box.h <= 0;
			if (nodes.box.hidden) continue;
			place(nodes.box, box.x, box.y, box.w, box.h);
			const { ox, oy } = g.origin(q);
			nodes.plane.style.transform = `translate3d(${-ox}px, ${-oy}px, 0)`;
			nodes.lines.render(this.metrics, rows, cols, gridlines);
			if (sheet)
				nodes.cells.render(
					buildItems({
						sheet,
						metrics: this.metrics,
						zoom,
						rows,
						cols,
						view: (r, c) => this.cellViewAt(r, c),
						measure: (text, font) => this.measurer.measure(text, font),
						hasFormula: (r, c) => getCell(sheet, r, c)?.formula !== undefined,
						fitNumber: (r, c, width, font) =>
							fitNumber(this.workbook(), sheet, r, c, width, (text) =>
								this.measurer.measure(text, font),
							),
					}),
					zoom,
				);
			nodes.overlay.render(this.metrics, {
				selection: selection.sheet === this.sheetIndex() ? selection : undefined,
				activeRange,
				copyRange: this.overlay.copyRange,
				references: this.overlay.references,
				fillPreview: this.overlay.fillPreview,
				showHandle: editable,
				maxRow,
				maxCol,
			});
		}
		paintChrome(
			{
				colHeaders: this.colHeaders,
				rowHeaders: this.rowHeaders,
				corner: this.corner,
				freezeH: this.#freezeH,
				freezeV: this.#freezeV,
			},
			g,
			this.metrics,
			vis,
			selection,
		);
		for (const hook of this.hooks) hook(this);
	}

	destroy(): void {
		this.#cancelFrame?.();
		this.root.remove();
		this.hooks.clear();
	}
}
