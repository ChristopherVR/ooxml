// The AutoFilter drop-down menu: sort, clear, colour, Text / Number / Date Filters, search and
// the value checklist. The criteria it builds are core's (`filter-menu.ts`); this paints them.
import {
	checklistCriteria,
	colorFilterCriteria,
	displayText,
	getCell,
	filterMenuModel,
	type FilterCriteria,
	type FilterMenuModel,
} from 'ooxml-core/xlsx';
import { h } from './dom';
import { presetItems, presetTitle, sortLabels, type PresetItem } from './filter-presets';
import type { GridView } from './grid-view';

const MAX_LISTED = 2000;

export interface FilterMenuHandle {
	close(): void;
}

interface Item {
	label: string;
	run?: () => void;
	disabled?: boolean;
	/** The label is already translated. */
	raw?: boolean;
	children?: HTMLElement;
	separatorBefore?: boolean;
}

export function openFilterMenu(
	view: GridView,
	col: number,
	anchor: { x: number; y: number; w: number; h: number },
	onClose: () => void,
	refocus: () => void,
): FilterMenuHandle | undefined {
	const ctx = view.ctx;
	const workbook = view.workbook();
	const sheet = view.sheet();
	const session = ctx.session();
	const filter = sheet?.autoFilter;
	if (!workbook || !sheet || !session || !filter) return undefined;
	const sheetIndex = view.sheetIndex();
	const offset = col - filter.range.start.col;
	const model = filterMenuModel(workbook, sheet, filter, offset);
	const header = displayText(workbook, getCell(sheet, filter.range.start.row, col));
	const name = header || ctx.t('Column {name}', { name: String(col + 1) });
	const doc = view.doc;
	const popup = h(doc, 'div', 'xg-popup xg-fmenu', {
		role: 'dialog',
		'aria-label': ctx.t('Filter menu'),
	});
	let closed = false;
	const dismiss = (event: Event) => {
		if (!event.composedPath().includes(popup)) close();
	};
	const onKey = (event: KeyboardEvent) => {
		if (event.key === 'Escape') {
			event.stopPropagation();
			close(true);
		}
	};
	const close = (focus = false) => {
		if (closed) return;
		closed = true;
		doc.removeEventListener('pointerdown', dismiss, true);
		popup.remove();
		onClose();
		if (focus) refocus();
	};
	const apply = (criteria: FilterCriteria | undefined) => {
		close(true);
		session.setFilterCriteria(sheetIndex, col, criteria);
	};
	const openDialog = (dialog: string, props: object) => {
		close();
		void ctx.dialogs.open(dialog, { col, kind: model.kind, ...props });
	};

	const itemButton = (item: Item): HTMLElement => {
		const wrap = h(doc, 'div', item.separatorBefore ? 'xg-fm-row xg-fm-sep' : 'xg-fm-row');
		const button = h(doc, 'button', 'xg-fm-item', { type: 'button', role: 'menuitem' });
		button.textContent = item.raw ? item.label : ctx.t(item.label);
		if (item.disabled) button.disabled = true;
		if (item.children) {
			button.setAttribute('aria-haspopup', 'menu');
			button.setAttribute('aria-expanded', 'false');
			button.classList.add('xg-fm-parent');
			const sub = item.children;
			sub.hidden = true;
			const toggle = (show: boolean) => {
				sub.hidden = !show;
				button.setAttribute('aria-expanded', String(show));
			};
			// Hover opens the submenu before the click lands, so a click opens it too, as in Excel.
			button.addEventListener('click', () => toggle(true));
			button.addEventListener('mouseenter', () => !item.disabled && toggle(true));
			wrap.addEventListener('mouseleave', () => toggle(false));
			button.addEventListener('keydown', (event) => {
				if (event.key === 'ArrowRight') toggle(true);
				else if (event.key === 'ArrowLeft') toggle(false);
			});
			wrap.append(button, sub);
		} else {
			button.addEventListener('click', () => item.run?.());
			wrap.append(button);
		}
		return wrap;
	};
	const submenu = (items: Item[]): HTMLElement => {
		const box = h(doc, 'div', 'xg-popup xg-fm-sub', { role: 'menu' });
		for (const item of items) box.append(itemButton(item));
		return box;
	};
	const presetMenu = (items: PresetItem[]): Item[] =>
		items.map((preset) => ({
			label: preset.label,
			...(preset.separatorBefore ? { separatorBefore: true } : {}),
			...(preset.children ? { children: submenu(presetMenu(preset.children)) } : {}),
			...(preset.dynamic
				? { run: () => apply({ dynamic: { type: preset.dynamic! } }) }
				: preset.top10
					? { run: () => openDialog('top10-filter', {}) }
					: preset.custom
						? { run: () => openDialog('custom-autofilter', preset.custom!) }
						: {}),
		}));

	const [asc, desc] = sortLabels(model.kind);
	const sort = (descending: boolean) => () => {
		close(true);
		session.sortByColumn(sheetIndex, col, descending);
	};
	const hasColors = model.fills.length > 0 || model.fonts.some((c) => c !== '#000000');
	const top: Item[] = [
		{ label: asc, run: sort(false) },
		{ label: desc, run: sort(true) },
		{
			label: ctx.t('Clear Filter From "{name}"', { name }),
			run: () => apply(undefined),
			disabled: !model.hasCriteria,
			raw: true,
			separatorBefore: true,
		},
		{
			label: 'Filter by Color',
			disabled: !hasColors,
			children: submenu([]),
		},
		{
			label: presetTitle(model.kind),
			children: submenu(presetMenu(presetItems(model.kind))),
		},
	];
	const colorHost = top[3]!.children!;
	colorHost.append(...colorSections(view, model, apply));
	for (const item of top) {
		popup.append(itemButton(item));
	}

	const checklist = checklistPanel(view, model, apply, close);
	popup.append(checklist);
	popup.addEventListener('keydown', onKey);
	popup.addEventListener('pointerdown', (event) => event.stopPropagation());
	const width = 250;
	const left = Math.max(
		view.geometry.headerW,
		Math.min(anchor.x + anchor.w - width, view.geometry.width - width - 4),
	);
	popup.style.left = `${left}px`;
	popup.style.top = `${anchor.y + anchor.h + 1}px`;
	popup.style.width = `${width}px`;
	view.view.append(popup);
	doc.addEventListener('pointerdown', dismiss, true);
	checklist.querySelector<HTMLElement>('input[type=search]')?.focus();
	return { close: () => close() };
}

function colorSections(
	view: GridView,
	model: FilterMenuModel,
	apply: (c: FilterCriteria) => void,
): HTMLElement[] {
	const { doc, ctx } = view;
	const section = (title: string, colors: string[], noFill: boolean, cell: boolean) => {
		const box = h(doc, 'div', 'xg-fm-colors', { role: 'group', 'aria-label': ctx.t(title) });
		const head = h(doc, 'div', 'xg-fm-head');
		head.textContent = ctx.t(title);
		box.append(head);
		const swatches = h(doc, 'div', 'xg-fm-swatches');
		for (const color of colors) {
			const b = h(doc, 'button', 'xg-fm-swatch', {
				type: 'button',
				'aria-label': color,
				title: color,
			});
			b.style.background = color;
			b.addEventListener('click', () => apply(colorFilterCriteria(color, cell)));
			swatches.append(b);
		}
		box.append(swatches);
		if (noFill) {
			const b = h(doc, 'button', 'xg-fm-item', { type: 'button', role: 'menuitem' });
			b.textContent = ctx.t(cell ? 'No Fill' : 'Automatic');
			b.addEventListener('click', () => apply(colorFilterCriteria(undefined, cell)));
			box.append(b);
		}
		return box;
	};
	const out: HTMLElement[] = [];
	if (model.fills.length || model.noFill)
		out.push(
			section('Filter by Cell Color', model.fills, model.noFill && model.fills.length > 0, true),
		);
	const fonts = model.fonts.filter((c) => c !== '#000000');
	if (fonts.length) out.push(section('Filter by Font Color', fonts, true, false));
	return out;
}

function checklistPanel(
	view: GridView,
	model: FilterMenuModel,
	apply: (c: FilterCriteria | undefined) => void,
	close: (refocus?: boolean) => void,
): HTMLElement {
	const { doc, ctx } = view;
	const panel = h(doc, 'div', 'xg-fm-values');
	const search = h(doc, 'input', 'xg-fm-search', {
		type: 'search',
		placeholder: ctx.t('Search'),
		'aria-label': ctx.t('Search'),
	});
	const list = h(doc, 'ul', 'xg-fm-list', { role: 'group', 'aria-label': ctx.t('Values') });
	const checked = new Set(model.values.filter((v) => v.checked).map((v) => v.text));
	let blanks = model.blanks?.checked ?? false;
	let query = '';
	const selectAll = h(doc, 'input', '', { type: 'checkbox', 'aria-label': ctx.t('(Select All)') });
	const visible = () =>
		model.values.filter((v) => !query || v.text.toLowerCase().includes(query)).slice(0, MAX_LISTED);
	const showBlanks = () => !!model.blanks && (!query || '(blanks)'.includes(query));
	const entry = (label: string, on: boolean, toggle: (on: boolean) => void) => {
		const li = h(doc, 'li');
		const row = h(doc, 'label', 'xg-fm-check');
		const box = h(doc, 'input', '', { type: 'checkbox' });
		box.checked = on;
		box.addEventListener('change', () => {
			toggle(box.checked);
			sync();
		});
		const text = h(doc, 'span');
		text.textContent = label;
		row.append(box, text);
		li.append(row);
		return li;
	};
	const sync = () => {
		const shown = visible();
		const total = shown.length + (showBlanks() ? 1 : 0);
		const on = shown.filter((v) => checked.has(v.text)).length + (showBlanks() && blanks ? 1 : 0);
		selectAll.checked = total > 0 && on === total;
		selectAll.indeterminate = on > 0 && on < total;
	};
	const render = () => {
		list.replaceChildren();
		for (const v of visible())
			list.append(
				entry(v.text, checked.has(v.text), (on) =>
					on ? checked.add(v.text) : checked.delete(v.text),
				),
			);
		if (showBlanks()) list.append(entry(ctx.t('(Blanks)'), blanks, (on) => (blanks = on)));
		if (!list.childElementCount) {
			const empty = h(doc, 'li', 'xg-fm-empty');
			empty.textContent = ctx.t('(No values)');
			list.append(empty);
		}
		sync();
	};
	selectAll.addEventListener('change', () => {
		for (const v of visible()) {
			if (selectAll.checked) checked.add(v.text);
			else checked.delete(v.text);
		}
		if (showBlanks()) blanks = selectAll.checked;
		render();
	});
	search.addEventListener('input', () => {
		query = search.value.trim().toLowerCase();
		// Like Excel: a search starts from "only the matches ticked".
		checked.clear();
		for (const v of visible()) checked.add(v.text);
		blanks = showBlanks();
		render();
	});
	const all = h(doc, 'label', 'xg-fm-check xg-fm-all');
	const allText = h(doc, 'span');
	allText.textContent = ctx.t('(Select All)');
	all.append(selectAll, allText);
	const buttons = h(doc, 'div', 'xg-fm-buttons');
	const ok = h(doc, 'button', 'xg-fm-ok', { type: 'button' });
	ok.textContent = ctx.t('OK');
	const cancel = h(doc, 'button', '', { type: 'button' });
	cancel.textContent = ctx.t('Cancel');
	ok.addEventListener('click', () => {
		// While searching only the ticked matches stay; otherwise every unticked value goes.
		apply(
			query
				? { values: [...checked], ...(blanks && showBlanks() ? { blank: true as const } : {}) }
				: checklistCriteria(model, checked, blanks),
		);
	});
	cancel.addEventListener('click', () => close(true));
	search.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') ok.click();
	});
	buttons.append(ok, cancel);
	panel.append(search, all, list, buttons);
	render();
	return panel;
}
