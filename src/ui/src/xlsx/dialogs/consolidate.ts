// Data > Consolidate: aggregate values across multiple source ranges into the active cell.
import {
	formatRange,
	type CellRange,
	type ConsolidateResult,
	type ConsolidateSource,
	type SubtotalFunction,
} from 'ooxml-core/xlsx';
import { target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import {
	checkbox,
	field,
	fieldset,
	invalid,
	listBox,
	row,
	select,
	text,
	textInput,
} from './fields';
import { button, showDialog } from './frame';
import { resolveReference } from './go-to';
import { SUBTOTAL_FUNCTION_OPTIONS } from './subtotal';

function formatSourceKey(sheetName: string, range: CellRange): string {
	const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)
		? sheetName
		: `'${sheetName.replace(/'/g, "''")}'`;
	return `${quoted}!${formatRange(range)}`;
}

export function openConsolidate(ctx: EditorContext): Promise<ConsolidateResult | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);

	const fn = select(ctx, SUBTOTAL_FUNCTION_OPTIONS, 'sum');
	const reference = textInput(ctx);
	const addBtn = button(ctx, 'Add');
	const delBtn = button(ctx, 'Delete');
	const sources: { key: string; source: ConsolidateSource }[] = [];
	const list = listBox(ctx, 'All references:');
	const topRow = checkbox(ctx, 'Top row', false);
	const leftCol = checkbox(ctx, 'Left column', false);

	const syncList = (): void => {
		list.setItems(sources.map((s) => [s.key, s.key] as const));
		delBtn.disabled = sources.length === 0;
	};
	const addRef = (raw: string): boolean => {
		const resolved = resolveReference(ctx, raw);
		if (!resolved) return false;
		const sheetName = t.workbook.sheets[resolved.sheet]?.name ?? `Sheet${resolved.sheet + 1}`;
		const key = formatSourceKey(sheetName, resolved.range);
		if (!sources.some((s) => s.key === key)) {
			sources.push({ key, source: { sheet: resolved.sheet, range: resolved.range } });
			syncList();
		}
		list.select(key);
		reference.value = '';
		return true;
	};

	addBtn.addEventListener('click', () => {
		if (!addRef(reference.value)) invalid(ctx, reference, 'The reference is not valid.');
	});
	delBtn.addEventListener('click', () => {
		const selected = list.value();
		if (!selected) return;
		const idx = sources.findIndex((s) => s.key === selected);
		if (idx >= 0) {
			sources.splice(idx, 1);
			syncList();
		}
	});
	syncList();

	return showDialog<ConsolidateResult>(ctx, {
		name: 'consolidate',
		heading: 'Consolidate',
		body: [
			field(ctx, 'Function:', fn),
			row(ctx, field(ctx, 'Reference:', reference), addBtn),
			text(ctx, 'All references:', 'xve-field-label'),
			row(ctx, list.element, delBtn),
			fieldset(ctx, 'Use labels in', topRow.wrapper, leftCol.wrapper),
		],
		opened: () => reference.focus(),
		submit: () => {
			if (reference.value.trim() && !addRef(reference.value))
				return invalid(ctx, reference, 'The reference is not valid.');
			if (!sources.length)
				return invalid(ctx, reference, 'Add at least one reference to consolidate.');
			return t.session.consolidate(
				t.sheet,
				t.active,
				sources.map((s) => s.source),
				{
					fn: fn.value as SubtotalFunction,
					useTopRow: topRow.input.checked,
					useLeftColumn: leftCol.input.checked,
				},
			);
		},
	});
}
