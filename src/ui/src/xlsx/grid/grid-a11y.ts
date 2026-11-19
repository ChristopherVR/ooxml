// Accessibility: a transparent grid of rows and cells over the painted window (the core's
// `accessibleWindow`), the focus sink's label and active descendant, and a polite live region
// announcing selection changes. The painted cells and headers are hidden from assistive
// technology; this layer is what a screen reader explores as a table.
import {
	accessibleWindow,
	cellKey,
	columnLabel,
	describeCell,
	formatAddress,
	MAX_COL,
	MAX_ROW,
	type AccessibleWindow,
	type CellRange,
} from 'ooxml-core/xlsx';
import type { Selection } from 'ooxml-core/xlsx/ui';
import { h } from './dom';
import type { GridView } from './grid-view';
import { selectionCellCount, selectionRef } from 'ooxml-core/xlsx/ui';

/** What the active cell announces: column context, reference, value, formula, comment, range. */
export function describeSelection(view: GridView, selection: Selection): string {
	const t = view.ctx.t;
	const { row, col } = selection.active;
	const sheet = view.sheet();
	const about = sheet ? describeCell(sheet, selection.active) : undefined;
	const ref = about?.ref ?? formatAddress(selection.active);
	const text = view.cellViewAt(row, col)?.text ?? '';
	let cell = text ? t('{ref}, {value}', { ref, value: text }) : t('{ref}, blank', { ref });
	if (about?.header) cell = t('{header}, {cell}', { header: about.header, cell });
	const parts = [cell];
	if (about?.formula) parts.push(t('Formula: {formula}', { formula: about.formula }));
	if (about?.comment) parts.push(t('Comment: {text}', { text: about.comment }));
	if (selectionCellCount(selection) > 1)
		parts.push(t('Selected {range}', { range: selectionRef(selection) }));
	return parts.join('. ');
}

const inRanges = (ranges: readonly CellRange[], row: number, col: number): boolean =>
	ranges.some(
		(r) => row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col,
	);

interface RowNode extends HTMLDivElement {
	xgSig?: string;
	xgCells?: Map<number, HTMLDivElement>;
}

interface RowCell {
	col: number;
	selected: boolean;
	text: string;
}

let instances = 0;
const SETTLE_MS = 120;

export class GridA11y {
	readonly live: HTMLDivElement;
	/** The accessible grid: a header row of column headers, then one row per sheet row. */
	readonly grid: HTMLDivElement;
	readonly input: HTMLTextAreaElement;
	#view: GridView;
	#editing: () => boolean;
	#timer: ReturnType<typeof setTimeout> | undefined;
	#prefix = `xg-a11y-${++instances}-`;
	#header: HTMLDivElement;
	#headerCols = '';
	#rows = new Map<number, RowNode>();
	#active: HTMLDivElement | undefined;
	#synced: Selection | undefined;
	#settle: ReturnType<typeof setTimeout> | undefined;
	/**
	 * A paint for a new selection syncs at once, so the active cell is pointed at in the same
	 * frame. Other paints (scrolling, model changes) sync once they stop for SETTLE_MS: rebuilding
	 * the layer on every scroll frame measurably slowed scrolling.
	 */
	#hook = () => {
		if (this.#settle) clearTimeout(this.#settle);
		if (this.#view.ctx.selection.get() !== this.#synced) this.sync();
		else this.#settle = setTimeout(() => this.sync(), SETTLE_MS);
	};

	constructor(view: GridView, input: HTMLTextAreaElement, editing: () => boolean) {
		this.#view = view;
		this.#editing = editing;
		this.input = input;
		const doc = view.doc;
		// Row and column indices count the header row and the row-header column, as ARIA requires.
		this.grid = h(doc, 'div', 'xg-a11y', {
			role: 'grid',
			'aria-label': view.ctx.t('Spreadsheet grid'),
			'aria-rowcount': String(MAX_ROW + 2),
			'aria-colcount': String(MAX_COL + 2),
			'aria-multiselectable': 'true',
		});
		this.#header = h(doc, 'div', '', { role: 'row', 'aria-rowindex': '1' });
		this.grid.append(this.#header);
		this.live = h(doc, 'div', 'xg-live', { 'aria-live': 'polite', role: 'status' });
		view.root.append(this.grid, this.live);
		view.hooks.add(this.#hook);
	}

	/** Labels the sink and announces (debounced) the new active cell after a selection change. */
	update(selection: Selection): void {
		const text = describeSelection(this.#view, selection);
		if (!this.#editing()) this.input.setAttribute('aria-label', text);
		if (this.#timer) clearTimeout(this.#timer);
		this.#timer = setTimeout(() => {
			this.live.textContent = text;
		}, 150);
	}

	/** Rebuilds the rows that changed after a paint, then points the sink at the active cell. */
	sync(): void {
		const view = this.#view;
		const selection = view.ctx.selection.get();
		this.#synced = selection;
		if (this.#settle) clearTimeout(this.#settle);
		this.#settle = undefined;
		const shown = selection.sheet === view.sheetIndex();
		const active = shown ? view.activeRange().start : undefined;
		const area = accessibleWindow(view.sheet(), view.visible, active);
		const colsKey = area.cols.join(',');
		if (colsKey !== this.#headerCols) {
			this.#headerCols = colsKey;
			this.#header.replaceChildren(
				...area.cols.map((col) => {
					const node = h(view.doc, 'div', '', {
						role: 'columnheader',
						'aria-colindex': String(col + 2),
					});
					node.textContent = columnLabel(col);
					return node;
				}),
			);
		}
		const keep = new Set(area.rows);
		for (const [row, node] of this.#rows)
			if (!keep.has(row)) {
				node.remove();
				this.#rows.delete(row);
			}
		const ranges = shown ? selection.ranges : [];
		let previous: Element = this.#header;
		for (const row of area.rows) {
			let node = this.#rows.get(row);
			if (!node) {
				node = h(view.doc, 'div', '', { role: 'row', 'aria-rowindex': String(row + 2) });
				this.#rows.set(row, node);
			}
			const cells: RowCell[] = [];
			for (const col of area.cols)
				if (!area.covered.has(cellKey(row, col)))
					cells.push({
						col,
						selected: inRanges(ranges, row, col),
						text: view.cellViewAt(row, col)?.text ?? '',
					});
			const sig = cells.map((c) => `${c.col}${c.selected ? '+' : '-'}${c.text}`).join('\u0000');
			if (node.xgSig !== sig) {
				node.xgSig = sig;
				this.#fillRow(node, row, area, cells);
			}
			if (previous.nextSibling !== node) previous.after(node);
			previous = node;
		}
		this.#point(active);
	}

	#fillRow(node: RowNode, row: number, area: AccessibleWindow, cells: RowCell[]): void {
		const doc = this.#view.doc;
		const head = h(doc, 'div', '', { role: 'rowheader', 'aria-colindex': '1' });
		head.textContent = String(row + 1);
		node.xgCells = new Map();
		for (const { col, selected, text } of cells) {
			const cell = h(doc, 'div', '', {
				role: 'gridcell',
				id: `${this.#prefix}${row}-${col}`,
				'aria-colindex': String(col + 2),
				'aria-selected': String(selected),
			});
			const span = area.spans.get(cellKey(row, col));
			if (span && span.rows > 1) cell.setAttribute('aria-rowspan', String(span.rows));
			if (span && span.cols > 1) cell.setAttribute('aria-colspan', String(span.cols));
			cell.textContent = text;
			node.xgCells.set(col, cell);
		}
		node.replaceChildren(head, ...node.xgCells.values());
	}

	/**
	 * Points the focus sink at the active cell's grid cell (while not editing, so typed text is
	 * still read from the editor), places that cell over the painted one for magnifiers and
	 * focus highlights, and refreshes the label, which may change without a selection change
	 * (the file opened, an undo, a recalculation).
	 */
	#point(active: { row: number; col: number } | undefined): void {
		const view = this.#view;
		const editing = this.#editing();
		const cell = active ? this.#rows.get(active.row)?.xgCells?.get(active.col) : undefined;
		if (this.#active !== cell) {
			this.#active?.removeAttribute('style');
			this.#active = cell;
		}
		if (cell && !editing && view.ctx.selection.get().drawing === undefined) {
			if (this.input.getAttribute('aria-activedescendant') !== cell.id)
				this.input.setAttribute('aria-activedescendant', cell.id);
			const box = view.geometry.rangeBox(view.activeRange());
			const g = view.geometry;
			if (box.x >= 0 && box.y >= 0 && box.x + box.w <= g.width && box.y + box.h <= g.height)
				cell.style.cssText = `left:${box.x}px;top:${box.y}px;width:${box.w}px;height:${box.h}px`;
			else cell.removeAttribute('style');
		} else this.input.removeAttribute('aria-activedescendant');
		if (!editing) {
			const label = describeSelection(view, view.ctx.selection.get());
			if (this.input.getAttribute('aria-label') !== label)
				this.input.setAttribute('aria-label', label);
		}
	}

	destroy(): void {
		if (this.#timer) clearTimeout(this.#timer);
		if (this.#settle) clearTimeout(this.#settle);
		this.#view.hooks.delete(this.#hook);
		this.grid.remove();
		this.live.remove();
	}
}
