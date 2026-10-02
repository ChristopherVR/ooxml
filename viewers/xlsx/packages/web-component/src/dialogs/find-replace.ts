// Find and Replace: Find / Replace tabs, options, Find All results, Find Next, Replace and
// Replace All, all through the core session's find API.
import { type FindMatch, type FindQuery, formatAddress, getCell } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { checkbox, el, field, row, select, tabs, textInput } from './fields.js';
import { button, showDialog } from './frame.js';
import { selectOn } from './go-to.js';

export interface FindReplaceProps {
	tab?: 'find' | 'replace';
}

/** Matches after the active cell first (wrapping), as Find Next walks them. */
export function orderFrom(
	matches: FindMatch[],
	sheet: number,
	row: number,
	col: number,
	byColumns: boolean,
): FindMatch[] {
	const after = (m: FindMatch): boolean =>
		m.sheet > sheet ||
		(m.sheet === sheet &&
			(byColumns
				? m.col > col || (m.col === col && m.row > row)
				: m.row > row || (m.row === row && m.col > col)));
	const index = matches.findIndex(after);
	return index < 0 ? matches : [...matches.slice(index), ...matches.slice(0, index)];
}

export function openFindReplace(
	ctx: EditorContext,
	props: FindReplaceProps = {},
): Promise<number | undefined> {
	const doc = ctx.host.ownerDocument;
	const what = textInput(ctx);
	const withText = textInput(ctx);
	const within = select(
		ctx,
		[
			['sheet', 'Sheet'],
			['workbook', 'Workbook'],
		],
		'sheet',
	);
	const order = select(
		ctx,
		[
			['rows', 'By Rows'],
			['columns', 'By Columns'],
		],
		'rows',
	);
	const lookIn = select(
		ctx,
		[
			['formulas', 'Formulas'],
			['values', 'Values'],
			['comments', 'Notes'],
		],
		'formulas',
	);
	const matchCase = checkbox(ctx, 'Match case');
	const whole = checkbox(ctx, 'Match entire cell contents');
	const options = el(ctx, 'div', 'xve-find-options');
	options.hidden = true;
	options.append(
		row(
			ctx,
			field(ctx, 'Within:', within),
			field(ctx, 'Search:', order),
			field(ctx, 'Look in:', lookIn),
		),
		row(ctx, matchCase.wrapper, whole.wrapper),
	);
	const toggle = button(ctx, 'Options >>');
	toggle.setAttribute('aria-expanded', 'false');
	toggle.addEventListener('click', () => {
		options.hidden = !options.hidden;
		toggle.setAttribute('aria-expanded', String(!options.hidden));
		toggle.textContent = ctx.t(options.hidden ? 'Options >>' : 'Options <<');
	});
	const replaceRow = field(ctx, 'Replace with:', withText);
	const findPanel = el(ctx, 'div');
	const replacePanel = el(ctx, 'div');
	const view = tabs(
		ctx,
		[
			{ id: 'find', label: 'Find', panel: findPanel },
			{ id: 'replace', label: 'Replace', panel: replacePanel },
		],
		props.tab === 'replace' && !ctx.readOnly() ? 'replace' : 'find',
	);
	const shared = el(ctx, 'div', 'xve-tabpanel');
	shared.append(field(ctx, 'Find what:', what), replaceRow, row(ctx, toggle), options);
	const results = el(ctx, 'div', 'xve-results');
	results.hidden = true;
	const summary = el(ctx, 'p', 'xve-note');
	summary.hidden = true;
	const findAll = button(ctx, 'Find All');
	const findNext = button(ctx, 'Find Next');
	const replaceOne = button(ctx, 'Replace');
	const replaceAll = button(ctx, 'Replace All');
	const sync = (): void => {
		const replacing = view.current() === 'replace';
		replaceRow.hidden = !replacing;
		replaceOne.hidden = !replacing;
		replaceAll.hidden = !replacing;
		const comments = lookIn.querySelector<HTMLOptionElement>('option[value="comments"]');
		const values = lookIn.querySelector<HTMLOptionElement>('option[value="values"]');
		if (comments) comments.disabled = replacing;
		if (values) values.disabled = replacing;
		if (replacing) lookIn.value = 'formulas';
	};
	const replaceTab = view.element.querySelector<HTMLButtonElement>('[data-tab="replace"]');
	if (replaceTab && ctx.readOnly()) replaceTab.disabled = true;
	for (const b of view.element.querySelectorAll('[role="tab"]')) b.addEventListener('click', sync);
	view.element.append(shared);
	sync();

	const query = (): FindQuery | undefined => {
		const session = ctx.session();
		if (!session || !what.value) return undefined;
		const q: FindQuery = {
			text: what.value,
			matchCase: matchCase.input.checked,
			wholeCell: whole.input.checked,
			lookIn: lookIn.value as NonNullable<FindQuery['lookIn']>,
			order: order.value as 'rows' | 'columns',
			wildcards: true,
		};
		if (within.value === 'sheet') q.sheet = ctx.activeSheet();
		return q;
	};
	let current: FindMatch | undefined;
	const show = (m: FindMatch): void => {
		current = m;
		const at = { row: m.row, col: m.col };
		selectOn(ctx, m.sheet, { start: at, end: at });
	};
	const notFound = (): void =>
		ctx.toast(ctx.t("We couldn't find what you were looking for."), 'info');
	const next = (): boolean => {
		const q = query();
		const session = ctx.session();
		if (!q || !session) return false;
		const active = ctx.selection.get().active;
		const list = orderFrom(
			session.findAll(q),
			ctx.activeSheet(),
			active.row,
			active.col,
			q.order === 'columns',
		);
		const m = list[0];
		if (!m) {
			notFound();
			return false;
		}
		show(m);
		return true;
	};
	findNext.addEventListener('click', () => void next());
	findAll.addEventListener('click', () => {
		const q = query();
		const session = ctx.session();
		if (!q || !session) return;
		const matches = session.findAll(q);
		const table = doc.createElement('table');
		const head = table.createTHead().insertRow();
		for (const label of ['Sheet', 'Cell', 'Value']) {
			const th = doc.createElement('th');
			th.scope = 'col';
			th.textContent = ctx.t(label);
			head.append(th);
		}
		const body = table.createTBody();
		for (const m of matches) {
			const tr = body.insertRow();
			tr.tabIndex = 0;
			const ws = session.workbook.sheets[m.sheet];
			const cell = ws ? getCell(ws, m.row, m.col) : undefined;
			for (const value of [
				ws?.name ?? '',
				formatAddress({ row: m.row, col: m.col }),
				cell?.formula !== undefined ? `=${cell.formula}` : m.text,
			])
				tr.insertCell().textContent = value;
			const pick = (): void => {
				for (const other of body.rows) other.removeAttribute('aria-selected');
				tr.setAttribute('aria-selected', 'true');
				show(m);
			};
			tr.addEventListener('click', pick);
			tr.addEventListener('keydown', (event) => {
				if (event.key === 'Enter') {
					event.preventDefault();
					event.stopPropagation();
					pick();
				}
			});
		}
		summary.textContent = ctx.t('{count} cell(s) found', { count: matches.length });
		summary.hidden = false;
		results.replaceChildren(table);
		results.hidden = false;
		if (!matches.length) notFound();
	});
	replaceOne.addEventListener('click', () => {
		const q = query();
		const session = ctx.session();
		if (!q || !session || ctx.readOnly()) return;
		if (current) session.replaceOne(q, withText.value, current);
		current = undefined;
		next();
	});
	let replaced: number | undefined;
	replaceAll.addEventListener('click', () => {
		const q = query();
		const session = ctx.session();
		if (!q || !session || ctx.readOnly()) return;
		const count = session.replaceAll(q, withText.value);
		replaced = (replaced ?? 0) + count;
		if (count) ctx.toast(ctx.t('All done. We made {count} replacements.', { count }), 'info');
		else notFound();
	});
	return showDialog<number>(
		ctx,
		{
			name: 'find-replace',
			heading: 'Find and Replace',
			body: [view.element, summary, results],
			okLabel: null,
			buttons: [replaceAll, replaceOne, findAll, findNext],
			wide: true,
			opened: () => what.focus(),
		},
		() => {
			// Enter in "Find what" runs Find Next (the dialog has no OK button).
			what.addEventListener('keydown', (event) => {
				if (event.key === 'Enter') {
					event.preventDefault();
					next();
				}
			});
		},
	).then((result) => result ?? replaced);
}
