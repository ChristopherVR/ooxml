// The list data-validation drop-down: the arrow button beside the active cell and the list popup
// (also used by "Pick From Drop-down List", which offers the column's own values like Excel).
import {
	getValue,
	isCellError,
	listValidationOptions,
	validationAt,
	type CellValue,
} from '@christophervr/xlsx-core';
import { h, place } from './dom.js';
import type { GridView } from './grid-view.js';

const textOf = (value: CellValue): string =>
	value === null
		? ''
		: typeof value === 'boolean'
			? value
				? 'TRUE'
				: 'FALSE'
			: isCellError(value)
				? value.error
				: String(value);

/** Unique texts above and below a cell in its column (contiguous block), for the pick list. */
export function columnPickList(view: GridView, row: number, col: number, limit = 500): string[] {
	const sheet = view.sheet();
	if (!sheet) return [];
	const seen = new Set<string>();
	for (const dir of [-1, 1]) {
		for (let r = row + dir, n = 0; r >= 0 && n < 5000; r += dir, n++) {
			const text = textOf(getValue(sheet, r, col));
			if (!text) break;
			seen.add(text);
			if (seen.size >= limit) break;
		}
	}
	return [...seen].sort((a, b) => a.localeCompare(b));
}

export class ValidationList {
	readonly arrow: HTMLDivElement;
	#view: GridView;
	#popup: HTMLDivElement | undefined;
	#onDone: () => void;

	constructor(view: GridView, onDone: () => void) {
		this.#view = view;
		this.#onDone = onDone;
		this.arrow = h(view.doc, 'div', 'xg-dd', {
			role: 'button',
			'aria-label': view.ctx.t('Open the list of values'),
		});
		this.arrow.hidden = true;
		this.arrow.addEventListener('pointerdown', (event) => {
			event.preventDefault();
			event.stopPropagation();
			if (this.#popup) this.close();
			else this.open(false);
		});
		view.view.append(this.arrow);
		view.hooks.add(() => this.position());
	}

	#options(): string[] | undefined {
		const view = this.#view;
		const workbook = view.workbook();
		const session = view.ctx.session();
		const { row, col } = view.ctx.selection.get().active;
		if (!workbook) return undefined;
		return listValidationOptions(workbook, view.sheetIndex(), row, col, {
			evaluate: (formula, at) => (session ? session.calc.evaluate(formula, at) : null),
		});
	}

	position(): void {
		const view = this.#view;
		const workbook = view.workbook();
		const { row, col } = view.ctx.selection.get().active;
		const rule = workbook ? validationAt(workbook, view.sheetIndex(), row, col) : undefined;
		const show = !view.ctx.readOnly() && rule?.type === 'list' && rule.showDropDown !== false;
		this.arrow.hidden = !show;
		if (!show) {
			this.close();
			return;
		}
		const box = view.geometry.rangeBox(view.activeRange());
		const size = Math.max(13, Math.round(17 * (view.metrics.zoom / 100)));
		place(this.arrow, box.x + box.w + 1, box.y + box.h - size, size, size);
	}

	get isOpen(): boolean {
		return Boolean(this.#popup);
	}

	/** Opens the list; `pick` falls back to the column's values when there is no list rule. */
	open(pick: boolean): boolean {
		const view = this.#view;
		if (view.ctx.readOnly()) return false;
		const { row, col } = view.ctx.selection.get().active;
		const items = this.#options() ?? (pick ? columnPickList(view, row, col) : undefined);
		if (!items) return false;
		this.close();
		const doc = view.doc;
		const popup = h(doc, 'div', 'xg-popup');
		const list = h(doc, 'ul', 'xg-list', {
			role: 'listbox',
			tabindex: '0',
			'aria-label': view.ctx.t('List of values'),
		});
		const sheet = view.sheet();
		const current = sheet ? textOf(getValue(sheet, row, col)) : '';
		let active = Math.max(0, items.indexOf(current));
		const options = items.map((item, index) => {
			const option = h(doc, 'li', '', { role: 'option', id: `xg-opt-${index}` });
			option.textContent = item;
			option.addEventListener('pointerdown', (event) => {
				event.preventDefault();
				this.#choose(item);
			});
			list.append(option);
			return option;
		});
		const mark = () => {
			options.forEach((o, i) => o.setAttribute('aria-selected', String(i === active)));
			list.setAttribute('aria-activedescendant', `xg-opt-${active}`);
			options[active]?.scrollIntoView?.({ block: 'nearest' });
		};
		mark();
		list.addEventListener('keydown', (event) => {
			event.stopPropagation();
			if (event.key === 'ArrowDown') active = Math.min(items.length - 1, active + 1);
			else if (event.key === 'ArrowUp') active = Math.max(0, active - 1);
			else if (event.key === 'Enter' || event.key === 'Tab') {
				const item = items[active];
				if (item !== undefined) this.#choose(item);
			} else if (event.key === 'Escape') this.close(true);
			else return;
			event.preventDefault();
			mark();
		});
		popup.append(list);
		const box = view.geometry.rangeBox(view.activeRange());
		popup.style.left = `${box.x}px`;
		popup.style.top = `${box.y + box.h + 1}px`;
		popup.style.minWidth = `${box.w + 18}px`;
		view.view.append(popup);
		this.#popup = popup;
		if (!items.length) {
			const empty = h(doc, 'li', '', { role: 'option', 'aria-disabled': 'true' });
			empty.textContent = view.ctx.t('(No values)');
			list.append(empty);
		}
		list.focus();
		return true;
	}

	#choose(value: string): void {
		const view = this.#view;
		const session = view.ctx.session();
		const { row, col } = view.ctx.selection.get().active;
		this.close(true);
		session?.setCellInput(view.sheetIndex(), row, col, value);
	}

	close(refocus = false): void {
		this.#popup?.remove();
		this.#popup = undefined;
		if (refocus) this.#onDone();
	}
}
