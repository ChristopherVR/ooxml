// The drop-down buttons in the header row of the sheet's AutoFilter, one per column, drawn over
// the cells like Excel's. A filtered column shows a funnel; the button opens the filter menu.
import { hasFilterCriteria } from 'ooxml-core/xlsx';
import { h, place, toggleClass } from './dom';
import { openFilterMenu, type FilterMenuHandle } from './filter-menu';
import type { GridView } from './grid-view';

export class FilterButtons {
	readonly #view: GridView;
	readonly #refocus: () => void;
	readonly #buttons = new Map<number, HTMLButtonElement>();
	#menu: FilterMenuHandle | undefined;
	#openCol = -1;

	constructor(view: GridView, refocus: () => void) {
		this.#view = view;
		this.#refocus = refocus;
		view.hooks.add(() => this.position());
	}

	get isOpen(): boolean {
		return !!this.#menu;
	}

	#button(col: number): HTMLButtonElement {
		let button = this.#buttons.get(col);
		if (button) return button;
		const view = this.#view;
		button = h(view.doc, 'button', 'xg-fb', { type: 'button', 'aria-haspopup': 'dialog' });
		button.addEventListener('pointerdown', (event) => {
			event.preventDefault();
			event.stopPropagation();
			if (this.#openCol === col) this.close();
			else this.open(col);
		});
		button.addEventListener('keydown', (event) => {
			if (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowDown') {
				event.preventDefault();
				event.stopPropagation();
				this.open(col);
			}
		});
		view.view.append(button);
		this.#buttons.set(col, button);
		return button;
	}

	/** Positions the buttons for the columns on screen and removes the rest. */
	position(): void {
		const view = this.#view;
		const sheet = view.sheet();
		const filter = sheet?.autoFilter;
		const g = view.geometry;
		if (!filter || view.ctx.readOnly()) {
			this.#clear();
			return;
		}
		const row = filter.range.start.row;
		const used = new Set<number>();
		const rowHidden = view.metrics.rowHeight(row) === 0;
		for (let col = filter.range.start.col; col <= filter.range.end.col; col++) {
			const column = filter.columns?.find((c) => c.offset === col - filter.range.start.col);
			if (rowHidden || column?.hiddenButton || view.metrics.colWidth(col) === 0) continue;
			const box = g.rangeBox({ start: { row, col }, end: { row, col } });
			const frozenCol = col < g.fCols;
			const frozenRow = row < g.fRows;
			const minX = g.headerW + (frozenCol ? 0 : g.frozenW);
			const minY = g.headerH + (frozenRow ? 0 : g.frozenH);
			if (box.x < minX || box.y < minY || box.x + box.w > g.width || box.y + box.h > g.height)
				continue;
			const size = Math.min(Math.max(13, Math.round(17 * (view.metrics.zoom / 100))), box.h, box.w);
			const button = this.#button(col);
			used.add(col);
			place(button, box.x + box.w - size - 1, box.y + box.h - size - 1, size, size);
			const active = hasFilterCriteria(column);
			toggleClass(button, 'xg-fb-active', active);
			button.setAttribute(
				'aria-label',
				view.ctx.t(active ? 'Filter drop-down (filtered)' : 'Filter drop-down'),
			);
		}
		for (const [col, button] of this.#buttons)
			if (!used.has(col)) {
				if (col === this.#openCol) this.close();
				button.remove();
				this.#buttons.delete(col);
			}
	}

	open(col: number): void {
		const view = this.#view;
		this.close();
		const filter = view.sheet()?.autoFilter;
		if (!filter) return;
		const box = view.geometry.rangeBox({
			start: { row: filter.range.start.row, col },
			end: { row: filter.range.start.row, col },
		});
		this.#openCol = col;
		this.#menu = openFilterMenu(
			view,
			col,
			box,
			() => {
				this.#menu = undefined;
				this.#openCol = -1;
			},
			this.#refocus,
		);
		if (!this.#menu) this.#openCol = -1;
	}

	close(): void {
		this.#menu?.close();
		this.#menu = undefined;
		this.#openCol = -1;
	}

	#clear(): void {
		this.close();
		for (const button of this.#buttons.values()) button.remove();
		this.#buttons.clear();
	}

	destroy(): void {
		this.#clear();
	}
}
