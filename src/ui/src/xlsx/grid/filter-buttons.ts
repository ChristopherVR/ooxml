// The drop-down buttons in the header row of the sheet's AutoFilter and every table header row,
// one per column, drawn over the cells like Excel's. A filtered column shows a funnel; the button
// opens the filter menu.
import {
	hasFilterCriteria,
	tableAutoFilter,
	tableFilterBounds,
	type AutoFilter,
} from 'ooxml-core/xlsx';
import { h, place, toggleClass } from './dom';
import { openFilterMenu, type FilterMenuHandle, type FilterTarget } from './filter-menu';
import type { GridView } from './grid-view';

interface ButtonTarget {
	key: string;
	row: number;
	col: number;
	target: FilterTarget;
}

export class FilterButtons {
	readonly #view: GridView;
	readonly #refocus: () => void;
	readonly #buttons = new Map<string, HTMLButtonElement>();
	readonly #targets = new Map<string, ButtonTarget>();
	#menu: FilterMenuHandle | undefined;
	#openKey = '';

	constructor(view: GridView, refocus: () => void) {
		this.#view = view;
		this.#refocus = refocus;
		view.hooks.add(() => this.position());
	}

	get isOpen(): boolean {
		return !!this.#menu;
	}

	#button(key: string): HTMLButtonElement {
		let button = this.#buttons.get(key);
		if (button) return button;
		const view = this.#view;
		button = h(view.doc, 'button', 'xg-fb', { type: 'button', 'aria-haspopup': 'dialog' });
		button.addEventListener('pointerdown', (event) => {
			event.preventDefault();
			event.stopPropagation();
			if (this.#openKey === key) this.close();
			else this.#openByKey(key);
		});
		button.addEventListener('keydown', (event) => {
			if (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowDown') {
				event.preventDefault();
				event.stopPropagation();
				this.#openByKey(key);
			}
		});
		view.view.append(button);
		this.#buttons.set(key, button);
		return button;
	}

	/** Positions the buttons for the sheet AutoFilter and table headers on screen, removing the rest. */
	position(): void {
		const view = this.#view;
		const sheet = view.sheet();
		const g = view.geometry;
		if (!sheet || view.ctx.readOnly()) {
			this.#clear();
			return;
		}
		const sources: FilterTarget[] = [];
		if (sheet.autoFilter) sources.push({ filter: sheet.autoFilter });
		for (const table of sheet.tables) {
			if (!table.headerRow) continue;
			const active = tableAutoFilter(table);
			const filter: AutoFilter = active ?? { range: tableFilterBounds(table) };
			sources.push({ filter, tableName: table.name });
		}
		if (!sources.length) {
			this.#clear();
			return;
		}
		const used = new Set<string>();
		this.#targets.clear();
		for (const target of sources) {
			const { filter } = target;
			const row = filter.range.start.row;
			const rowHidden = view.metrics.rowHeight(row) === 0;
			if (rowHidden) continue;
			for (let col = filter.range.start.col; col <= filter.range.end.col; col++) {
				const key = `${row}:${col}`;
				if (used.has(key)) continue;
				const column = filter.columns?.find((c) => c.offset === col - filter.range.start.col);
				if (column?.hiddenButton || view.metrics.colWidth(col) === 0) continue;
				const box = g.rangeBox({ start: { row, col }, end: { row, col } });
				const frozenCol = col < g.fCols;
				const frozenRow = row < g.fRows;
				const minX = g.headerW + (frozenCol ? 0 : g.frozenW);
				const minY = g.headerH + (frozenRow ? 0 : g.frozenH);
				if (box.x < minX || box.y < minY || box.x + box.w > g.width || box.y + box.h > g.height)
					continue;
				const size = Math.min(
					Math.max(13, Math.round(17 * (view.metrics.zoom / 100))),
					box.h,
					box.w,
				);
				const button = this.#button(key);
				used.add(key);
				this.#targets.set(key, { key, row, col, target });
				place(button, box.x + box.w - size - 1, box.y + box.h - size - 1, size, size);
				const active = hasFilterCriteria(column);
				toggleClass(button, 'xg-fb-active', active);
				button.setAttribute(
					'aria-label',
					view.ctx.t(active ? 'Filter drop-down (filtered)' : 'Filter drop-down'),
				);
			}
		}
		for (const [key, button] of this.#buttons)
			if (!used.has(key)) {
				if (key === this.#openKey) this.close();
				button.remove();
				this.#buttons.delete(key);
			}
	}

	#openByKey(key: string): void {
		const entry = this.#targets.get(key);
		if (!entry) return;
		this.openTarget(entry.row, entry.col, entry.target);
	}

	/** Opens the sheet AutoFilter drop-down for `col` (or the first table covering `col`). */
	open(col: number): void {
		const sheet = this.#view.sheet();
		if (!sheet) return;
		if (
			sheet.autoFilter &&
			col >= sheet.autoFilter.range.start.col &&
			col <= sheet.autoFilter.range.end.col
		) {
			this.openTarget(sheet.autoFilter.range.start.row, col, { filter: sheet.autoFilter });
			return;
		}
		const table = sheet.tables.find(
			(t) => t.headerRow && col >= t.range.start.col && col <= t.range.end.col,
		);
		if (table) {
			const filter: AutoFilter = tableAutoFilter(table) ?? { range: tableFilterBounds(table) };
			this.openTarget(table.range.start.row, col, { filter, tableName: table.name });
		}
	}

	openTarget(row: number, col: number, target: FilterTarget): void {
		const view = this.#view;
		this.close();
		const box = view.geometry.rangeBox({
			start: { row, col },
			end: { row, col },
		});
		const key = `${row}:${col}`;
		this.#openKey = key;
		this.#menu = openFilterMenu(
			view,
			col,
			box,
			() => {
				this.#menu = undefined;
				this.#openKey = '';
			},
			this.#refocus,
			target,
		);
		if (!this.#menu) this.#openKey = '';
	}

	close(): void {
		this.#menu?.close();
		this.#menu = undefined;
		this.#openKey = '';
	}

	#clear(): void {
		this.close();
		for (const button of this.#buttons.values()) button.remove();
		this.#buttons.clear();
		this.#targets.clear();
	}

	destroy(): void {
		this.#clear();
	}
}
