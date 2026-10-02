// Accessibility: the grid role and counts, the focus sink's label (active cell reference and
// value) and a polite live region announcing selection changes.
import { formatAddress, MAX_COL, MAX_ROW } from '@christophervr/xlsx-core';
import type { Selection } from '../context.js';
import { h } from './dom.js';
import type { GridView } from './grid-view.js';
import { selectionCellCount, selectionRef } from './selection-ops.js';

export function describeSelection(view: GridView, selection: Selection): string {
	const t = view.ctx.t;
	const ref = formatAddress(selection.active);
	const cell = view.cellViewAt(selection.active.row, selection.active.col);
	const text = cell?.text ?? '';
	const base = text ? t('{ref}, {value}', { ref, value: text }) : t('{ref}, blank', { ref });
	const count = selectionCellCount(selection);
	if (count <= 1) return base;
	return `${base}. ${t('Selected {range}', { range: selectionRef(selection) })}`;
}

export class GridA11y {
	readonly live: HTMLDivElement;
	#view: GridView;
	#timer: ReturnType<typeof setTimeout> | undefined;

	constructor(view: GridView, input: HTMLTextAreaElement) {
		this.#view = view;
		const root = view.root;
		root.setAttribute('aria-rowcount', String(MAX_ROW + 1));
		root.setAttribute('aria-colcount', String(MAX_COL + 1));
		root.setAttribute('aria-multiselectable', 'true');
		root.setAttribute('aria-label', view.ctx.t('Spreadsheet grid'));
		this.live = h(view.doc, 'div', 'xg-live', { 'aria-live': 'polite', role: 'status' });
		root.append(this.live);
		this.input = input;
	}

	readonly input: HTMLTextAreaElement;

	/** Updates the sink label and announces (debounced) after a selection change. */
	update(selection: Selection, editing: boolean): void {
		const text = describeSelection(this.#view, selection);
		if (!editing) this.input.setAttribute('aria-label', text);
		if (this.#timer) clearTimeout(this.#timer);
		this.#timer = setTimeout(() => {
			this.live.textContent = text;
		}, 150);
	}

	destroy(): void {
		if (this.#timer) clearTimeout(this.#timer);
		this.live.remove();
	}
}
