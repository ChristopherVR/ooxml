// Name Manager: the workbook's defined names with value, reference, scope and comment, a filter,
// and New / Edit / Delete.
import { type CellValue, type DefinedName, isCellError } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { el, field, select } from './fields.js';
import { button, showDialog } from './frame.js';

type NameFilter = 'all' | 'worksheet' | 'workbook' | 'errors' | 'noErrors';

const FILTERS: ReadonlyArray<readonly [NameFilter, string]> = [
	['all', 'All'],
	['worksheet', 'Names Scoped to Worksheet'],
	['workbook', 'Names Scoped to Workbook'],
	['errors', 'Names with Errors'],
	['noErrors', 'Names without Errors'],
];

const valueText = (value: CellValue | undefined): string => {
	if (value === null || value === undefined) return '';
	if (isCellError(value)) return value.error;
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	return String(value);
};

/** A name's current value as the Name Manager shows it (errors as `#REF!` etc.). */
export function nameValue(ctx: EditorContext, name: DefinedName): { text: string; error: boolean } {
	const session = ctx.session();
	if (!session) return { text: '', error: false };
	try {
		const value = session.calc.evaluate(name.formula, {
			sheet: name.localSheet ?? ctx.activeSheet(),
			row: 0,
			col: 0,
		});
		return { text: valueText(value), error: isCellError(value) };
	} catch {
		return { text: '#NAME?', error: true };
	}
}

export const visibleNames = (names: readonly DefinedName[]): DefinedName[] =>
	names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.'));

export function openNameManager(ctx: EditorContext): Promise<undefined> {
	const doc = ctx.host.ownerDocument;
	const filter = select(ctx, FILTERS, 'all');
	const box = el(ctx, 'div', 'xve-results');
	const table = doc.createElement('table');
	box.append(table);
	const newButton = button(ctx, 'New...');
	const editButton = button(ctx, 'Edit...');
	const deleteButton = button(ctx, 'Delete');
	let selected: DefinedName | undefined;
	const sync = (): void => {
		const editable = !!selected && !ctx.readOnly();
		editButton.disabled = !editable;
		deleteButton.disabled = !editable;
		newButton.disabled = ctx.readOnly();
	};
	const render = (): void => {
		const workbook = ctx.workbook();
		table.replaceChildren();
		const head = table.createTHead().insertRow();
		for (const label of ['Name', 'Value', 'Refers To', 'Scope', 'Comment']) {
			const th = doc.createElement('th');
			th.scope = 'col';
			th.textContent = ctx.t(label);
			head.append(th);
		}
		const body = table.createTBody();
		const names = visibleNames(workbook?.definedNames ?? []);
		if (selected && !names.includes(selected)) selected = undefined;
		for (const name of names) {
			const value = nameValue(ctx, name);
			const kind = filter.value as NameFilter;
			if (kind === 'worksheet' && name.localSheet === undefined) continue;
			if (kind === 'workbook' && name.localSheet !== undefined) continue;
			if (kind === 'errors' && !value.error) continue;
			if (kind === 'noErrors' && value.error) continue;
			const tr = body.insertRow();
			tr.tabIndex = 0;
			tr.dataset.name = name.name;
			if (name === selected) tr.setAttribute('aria-selected', 'true');
			const scope =
				name.localSheet === undefined
					? ctx.t('Workbook')
					: (workbook?.sheets[name.localSheet]?.name ?? '');
			for (const text of [name.name, value.text, `=${name.formula}`, scope, name.comment ?? ''])
				tr.insertCell().textContent = text;
			const pick = (): void => {
				selected = name;
				for (const other of body.rows) other.removeAttribute('aria-selected');
				tr.setAttribute('aria-selected', 'true');
				sync();
			};
			tr.addEventListener('click', pick);
			tr.addEventListener('focus', pick);
			tr.addEventListener('dblclick', () => editButton.click());
			tr.addEventListener('keydown', (event) => {
				const to =
					event.key === 'ArrowDown'
						? tr.nextElementSibling
						: event.key === 'ArrowUp'
							? tr.previousElementSibling
							: null;
				if (to instanceof doc.defaultView!.HTMLElement) {
					event.preventDefault();
					to.focus();
				} else if (event.key === 'Delete') deleteButton.click();
			});
		}
		sync();
	};
	filter.addEventListener('change', render);
	newButton.addEventListener('click', async () => {
		await ctx.dialogs.open('define-name');
		render();
	});
	editButton.addEventListener('click', async () => {
		if (!selected) return;
		const props: { name: string; localSheet?: number } = { name: selected.name };
		if (selected.localSheet !== undefined) props.localSheet = selected.localSheet;
		await ctx.dialogs.open('define-name', props);
		render();
	});
	deleteButton.addEventListener('click', () => {
		const session = ctx.session();
		if (!selected || !session || ctx.readOnly()) return;
		session.deleteDefinedName(selected.name, selected.localSheet);
		selected = undefined;
		render();
	});
	render();
	return showDialog<undefined>(ctx, {
		name: 'name-manager',
		heading: 'Name Manager',
		wide: true,
		okLabel: null,
		buttons: [newButton, editButton, deleteButton],
		body: [field(ctx, 'Filter', filter), box],
		opened: () => (table.tBodies[0]?.rows[0] ?? newButton).focus(),
	});
}
