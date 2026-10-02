// Data > Sort: several sort levels (column, sort on cell values, order) over the selection or
// its current region, with the "My data has headers" switch.
import { type CellRange, columnLabel, getCell } from '@christophervr/xlsx-core';
import { guessHeader, regionOf, target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { checkbox, el, row, select } from './fields.js';
import { button, showDialog } from './frame.js';

export interface SortKeySpec {
	col: number;
	descending: boolean;
}

interface Level {
	element: HTMLElement;
	column: HTMLSelectElement;
	order: HTMLSelectElement;
}

export function sortDialog(ctx: EditorContext): Promise<SortKeySpec[] | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const region: CellRange = regionOf(t);
	const header = checkbox(ctx, 'My data has headers', guessHeader(t.ws, region));
	const levels: Level[] = [];
	const list = el(ctx, 'div', 'xve-sort-levels');
	let selected = 0;
	const columnOptions = (): [string, string][] => {
		const out: [string, string][] = [];
		for (let col = region.start.col; col <= region.end.col; col++) {
			const head = getCell(t.ws, region.start.row, col)?.value;
			const label =
				header.input.checked && head !== null && head !== undefined && String(head) !== ''
					? String(head)
					: ctx.t('Column {name}', { name: columnLabel(col) });
			out.push([String(col), label]);
		}
		return out;
	};
	const render = (): void => {
		list.replaceChildren(
			...levels.map((level, i) => {
				level.element.setAttribute('aria-selected', String(i === selected));
				const caption = level.element.querySelector('.xve-sort-caption');
				if (caption) caption.textContent = ctx.t(i === 0 ? 'Sort by' : 'Then by');
				return level.element;
			}),
		);
	};
	const addLevel = (col: number, descending: boolean, at = levels.length): void => {
		const column = select(ctx, columnOptions(), String(col), true);
		column.setAttribute('aria-label', ctx.t('Column'));
		const on = select(ctx, [['values', 'Cell Values']], 'values');
		on.setAttribute('aria-label', ctx.t('Sort On'));
		const order = select(
			ctx,
			[
				['asc', 'A to Z'],
				['desc', 'Z to A'],
			],
			descending ? 'desc' : 'asc',
		);
		order.setAttribute('aria-label', ctx.t('Order'));
		const caption = el(ctx, 'span', 'xve-field-label xve-sort-caption');
		const element = row(ctx, caption, column, on, order);
		const level: Level = { element, column, order };
		element.addEventListener('focusin', () => {
			selected = levels.indexOf(level);
			render();
		});
		levels.splice(at, 0, level);
		selected = at;
		render();
	};
	addLevel(
		t.active.col >= region.start.col && t.active.col <= region.end.col
			? t.active.col
			: region.start.col,
		false,
	);
	header.input.addEventListener('change', () => {
		const options = columnOptions();
		for (const level of levels) {
			const value = level.column.value;
			level.column.replaceChildren(...options.map(([v, label]) => new Option(label, v)));
			level.column.value = value;
		}
	});
	const add = button(ctx, 'Add Level');
	const remove = button(ctx, 'Delete Level');
	const copy = button(ctx, 'Copy Level');
	const up = button(ctx, 'Move Up');
	const down = button(ctx, 'Move Down');
	add.addEventListener('click', () => {
		const used = new Set(levels.map((l) => Number(l.column.value)));
		let col = region.start.col;
		while (used.has(col) && col < region.end.col) col++;
		addLevel(col, false);
	});
	remove.addEventListener('click', () => {
		if (levels.length <= 1) return;
		levels.splice(selected, 1);
		selected = Math.max(0, Math.min(selected, levels.length - 1));
		render();
	});
	copy.addEventListener('click', () => {
		const level = levels[selected];
		if (level) addLevel(Number(level.column.value), level.order.value === 'desc', selected + 1);
	});
	const move = (delta: number): void => {
		const to = selected + delta;
		if (to < 0 || to >= levels.length) return;
		const [level] = levels.splice(selected, 1);
		if (level) levels.splice(to, 0, level);
		selected = to;
		render();
	};
	up.addEventListener('click', () => move(-1));
	down.addEventListener('click', () => move(1));
	return showDialog<SortKeySpec[]>(ctx, {
		name: 'sort',
		heading: 'Sort',
		wide: true,
		body: [row(ctx, add, remove, copy, up, down, header.wrapper), list],
		opened: () => levels[0]?.column.focus(),
		submit: () => {
			const keys = levels.map((l) => ({
				col: Number(l.column.value),
				descending: l.order.value === 'desc',
			}));
			t.session.sort(t.sheet, region, keys, header.input.checked);
			return keys;
		},
	});
}
