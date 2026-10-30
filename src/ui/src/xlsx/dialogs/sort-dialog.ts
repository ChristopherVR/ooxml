// Data > Sort: several sort levels over the selection or its current region. Each level has a
// column (a row when sorting left to right), Sort On (cell values, cell or font colour, or a
// conditional-format icon) and an order (A to Z / Z to A / a custom list, or a colour or icon
// with On Top / On Bottom). Options... holds case sensitivity and the orientation.
import {
	BUILT_IN_SORT_LISTS,
	type CellRange,
	columnLabel,
	getCell,
	parseSortList,
	sortChoices,
	type SortChoices,
	type SortKey,
} from 'ooxml-core/xlsx';
import { guessHeader, regionOf, target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, el, row, select, textInput } from './fields';
import { button, showDialog } from './frame';

export type SortKeySpec = SortKey & { descending: boolean };

interface Level {
	element: HTMLElement;
	column: HTMLSelectElement;
	on: HTMLSelectElement;
	order: HTMLSelectElement;
	choice: HTMLSelectElement;
	listPick: HTMLSelectElement;
	listText: HTMLInputElement;
}

const NEW_LIST = 'new';

export function sortDialog(ctx: EditorContext): Promise<SortKeySpec[] | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const region: CellRange = regionOf(t);
	const header = checkbox(ctx, 'My data has headers', guessHeader(t.ws, region));
	const caseBox = checkbox(ctx, 'Case sensitive');
	const orientation = select(
		ctx,
		[
			['tb', 'Sort top to bottom'],
			['lr', 'Sort left to right'],
		],
		'tb',
	);
	orientation.setAttribute('aria-label', ctx.t('Orientation'));
	const optionsPanel = row(ctx, caseBox.wrapper, orientation);
	optionsPanel.hidden = true;
	const levels: Level[] = [];
	const list = el(ctx, 'div', 'xve-sort-levels');
	let selected = 0;
	const across = (): boolean => orientation.value === 'lr';
	const choicesCache = new Map<string, SortChoices>();
	const choicesFor = (key: number): SortChoices => {
		const id = `${across()}:${header.input.checked}:${key}`;
		let found = choicesCache.get(id);
		if (!found) {
			const workbook = ctx.workbook();
			found = workbook
				? sortChoices(
						workbook,
						t.sheet,
						(formula, at) => t.session.calc.evaluate(formula, at),
						region,
						key,
						header.input.checked,
						across(),
					)
				: { fills: [], noFill: false, fonts: [], icons: [], noIcon: false };
			choicesCache.set(id, found);
		}
		return found;
	};
	const lineOptions = (): [string, string][] => {
		const out: [string, string][] = [];
		const [from, to] = across()
			? [region.start.row, region.end.row]
			: [region.start.col, region.end.col];
		for (let at = from; at <= to; at++) {
			const head = across()
				? getCell(t.ws, at, region.start.col)?.value
				: getCell(t.ws, region.start.row, at)?.value;
			const named = header.input.checked && head !== null && head !== undefined && head !== '';
			out.push([
				String(at),
				named
					? String(head)
					: across()
						? ctx.t('Row {name}', { name: String(at + 1) })
						: ctx.t('Column {name}', { name: columnLabel(at) }),
			]);
		}
		return out;
	};
	const fill = (sel: HTMLSelectElement, options: [string, string][], raw = false): void => {
		const value = sel.value;
		sel.replaceChildren(...options.map(([v, label]) => new Option(raw ? label : ctx.t(label), v)));
		if (options.some(([v]) => v === value)) sel.value = value;
	};
	const refresh = (level: Level): void => {
		const on = level.on.value;
		const plain = on === 'value';
		const [before, after] = across() ? ['On Left', 'On Right'] : ['On Top', 'On Bottom'];
		fill(
			level.order,
			plain
				? [
						['asc', 'A to Z'],
						['desc', 'Z to A'],
						['custom', 'Custom List...'],
					]
				: [
						['top', before!],
						['bottom', after!],
					],
		);
		level.choice.hidden = plain;
		const found = choicesFor(Number(level.column.value));
		if (on === 'cellColor')
			fill(
				level.choice,
				[
					...(found.noFill ? [['', ctx.t('No Cell Color')] as [string, string]] : []),
					...found.fills.map((c): [string, string] => [c, c]),
				],
				true,
			);
		else if (on === 'fontColor')
			fill(
				level.choice,
				found.fonts.map((c): [string, string] => [c, c === '#000000' ? ctx.t('Automatic') : c]),
				true,
			);
		else if (on === 'icon')
			fill(
				level.choice,
				[
					...(found.noIcon ? [['none', ctx.t('No Cell Icon')] as [string, string]] : []),
					...found.icons.map((i): [string, string] => [
						`${i.set}:${i.index}`,
						`${i.set} ${i.index + 1}`,
					]),
				],
				true,
			);
		const custom = plain && level.order.value === 'custom';
		level.listPick.hidden = !custom;
		level.listText.hidden = !custom || level.listPick.value !== NEW_LIST;
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
	const addLevel = (col: number, descending: boolean, at = levels.length, from?: Level): void => {
		const column = select(ctx, lineOptions(), String(col), true);
		column.setAttribute('aria-label', ctx.t(across() ? 'Row' : 'Column'));
		const on = select(
			ctx,
			[
				['value', 'Cell Values'],
				['cellColor', 'Cell Color'],
				['fontColor', 'Font Color'],
				['icon', 'Conditional Formatting Icon'],
			],
			from?.on.value ?? 'value',
		);
		on.setAttribute('aria-label', ctx.t('Sort On'));
		const order = select(ctx, [['asc', 'A to Z']], descending ? 'desc' : 'asc');
		order.setAttribute('aria-label', ctx.t('Order'));
		const choice = select(ctx, [], undefined, true);
		choice.setAttribute('aria-label', ctx.t('Color'));
		const listPick = select(
			ctx,
			[
				...BUILT_IN_SORT_LISTS.map((items, i): [string, string] => [String(i), items.join(', ')]),
				[NEW_LIST, ctx.t('New list...')],
			],
			undefined,
			true,
		);
		listPick.setAttribute('aria-label', ctx.t('Custom List...'));
		const listText = textInput(ctx, '');
		listText.setAttribute('aria-label', ctx.t('List entries'));
		const caption = el(ctx, 'span', 'xve-field-label xve-sort-caption');
		const element = row(ctx, caption, column, on, choice, order, listPick, listText);
		const level: Level = { element, column, on, order, choice, listPick, listText };
		element.addEventListener('focusin', () => {
			selected = levels.indexOf(level);
			render();
		});
		for (const sel of [on, column, order, listPick])
			sel.addEventListener('change', () => refresh(level));
		levels.splice(at, 0, level);
		selected = at;
		refresh(level);
		if (from) {
			order.value = from.order.value;
			choice.value = from.choice.value;
			listPick.value = from.listPick.value;
			listText.value = from.listText.value;
			refresh(level);
		}
		render();
	};
	const rebuildColumns = (): void => {
		choicesCache.clear();
		const options = lineOptions();
		for (const level of levels) {
			const value = options.some(([v]) => v === level.column.value)
				? level.column.value
				: options[0]![0];
			fill(level.column, options, true);
			level.column.value = value;
			level.column.setAttribute('aria-label', ctx.t(across() ? 'Row' : 'Column'));
			refresh(level);
		}
	};
	addLevel(
		t.active.col >= region.start.col && t.active.col <= region.end.col
			? t.active.col
			: region.start.col,
		false,
	);
	header.input.addEventListener('change', rebuildColumns);
	orientation.addEventListener('change', () => {
		for (const level of levels) level.column.replaceChildren();
		rebuildColumns();
	});
	const add = button(ctx, 'Add Level');
	const remove = button(ctx, 'Delete Level');
	const copy = button(ctx, 'Copy Level');
	const up = button(ctx, 'Move Up');
	const down = button(ctx, 'Move Down');
	const options = button(ctx, 'Options...');
	options.addEventListener('click', () => (optionsPanel.hidden = !optionsPanel.hidden));
	add.addEventListener('click', () => {
		const used = new Set(levels.map((l) => Number(l.column.value)));
		const [from, to] = across()
			? [region.start.row, region.end.row]
			: [region.start.col, region.end.col];
		let at = from;
		while (used.has(at) && at < to) at++;
		addLevel(at, false);
	});
	remove.addEventListener('click', () => {
		if (levels.length <= 1) return;
		levels.splice(selected, 1);
		selected = Math.max(0, Math.min(selected, levels.length - 1));
		render();
	});
	copy.addEventListener('click', () => {
		const level = levels[selected];
		if (level)
			addLevel(Number(level.column.value), level.order.value === 'desc', selected + 1, level);
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
	const keyOf = (level: Level): SortKey & { descending: boolean } => {
		const col = Number(level.column.value);
		const on = level.on.value;
		if (on === 'value') {
			const spec: SortKeySpec = { col, descending: level.order.value === 'desc' };
			if (level.order.value === 'custom') {
				const text = level.listPick.value === NEW_LIST ? level.listText.value : '';
				const items =
					level.listPick.value === NEW_LIST
						? parseSortList(text)
						: [...(BUILT_IN_SORT_LISTS[Number(level.listPick.value)] ?? [])];
				if (items.length) spec.customList = items;
			}
			return spec;
		}
		const spec: SortKeySpec = { col, descending: level.order.value === 'bottom' };
		if (on === 'icon') {
			const [set, index] = level.choice.value.split(':');
			spec.sortOn = 'icon';
			spec.icon = set && index !== undefined ? { set, index: Number(index) } : { set: '' };
		} else {
			spec.sortOn = on === 'fontColor' ? 'fontColor' : 'cellColor';
			if (level.choice.value) spec.color = level.choice.value;
		}
		return spec;
	};
	return showDialog<SortKeySpec[]>(ctx, {
		name: 'sort',
		heading: 'Sort',
		wide: true,
		body: [row(ctx, add, remove, copy, up, down, options, header.wrapper), optionsPanel, list],
		opened: () => levels[0]?.column.focus(),
		submit: () => {
			const keys = levels.map(keyOf);
			t.session.sort(t.sheet, region, keys, header.input.checked, {
				...(across() ? { leftToRight: true } : {}),
				...(caseBox.input.checked ? { caseSensitive: true } : {}),
			});
			return keys;
		},
	});
}
