import {
	pivotPageItems,
	pivotPageItemLabel,
	type PivotPageFilter,
	type PivotTable,
	type CellValue,
} from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { field, select } from './fields';
export function pivotPageControls(
	ctx: EditorContext,
	initial: PivotPageFilter | undefined,
	data: () => Pick<PivotTable, 'fields' | 'records'>,
) {
	const source = select(ctx, [['', '(None)']], initial?.field ?? '');
	const item = select(ctx, [['', '(All)']], '');
	let values: CellValue[] = [];
	let loaded = false;
	let recordSource: CellValue[][] | undefined;
	let itemField: string | undefined;
	const loadItems = (chosen?: CellValue) => {
		const current = data().records;
		const changed = current !== recordSource || source.value !== itemField;
		if (changed) {
			recordSource = current;
			itemField = source.value;
			values = source.value ? pivotPageItems(data(), source.value) : [];
			item.replaceChildren();
			for (const [value, label, placeholder] of [
				['', '(All)', true] as const,
				...values.map(
					(v, i) => [String(i), String(pivotPageItemLabel(v)), v === null || v === ''] as const,
				),
			]) {
				const option = ctx.host.ownerDocument.createElement('option');
				option.value = value;
				option.textContent = placeholder ? ctx.t(label) : label;
				item.append(option);
			}
		}
		item.value =
			chosen === undefined
				? ''
				: String(values.findIndex((v) => JSON.stringify(v) === JSON.stringify(chosen)));
		if (item.selectedIndex < 0) item.value = '';
		item.disabled = !source.value;
	};
	source.addEventListener('change', () => loadItems());
	return {
		source,
		item,
		body: [field(ctx, 'Report Filter Field:', source), field(ctx, 'Report Filter Item:', item)],
		update: (excluded: string[]) => {
			const old = loaded ? source.value : (initial?.field ?? '');
			const fields = data().fields;
			source.replaceChildren();
			for (const name of ['', ...fields.filter((name) => !excluded.includes(name))]) {
				const option = ctx.host.ownerDocument.createElement('option');
				option.value = name;
				option.textContent = name || ctx.t('(None)');
				source.append(option);
			}
			source.value = old;
			if (source.selectedIndex < 0) source.value = '';
			const chosen = !loaded
				? initial?.item
				: source.value === old && item.value !== ''
					? values[Number(item.value)]
					: undefined;
			loadItems(chosen);
			loaded = true;
			item.disabled = !source.value;
		},
		read: (): PivotPageFilter | null =>
			source.value
				? {
						field: source.value,
						...(item.value !== '' ? { item: values[Number(item.value)]! } : {}),
					}
				: null,
	};
}
