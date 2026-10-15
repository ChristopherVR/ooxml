import {
	defaultFieldGroup,
	pivotGroupableFields,
	pivotTableAt,
	type PivotFieldGroup,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { field, numberInput, select } from './fields';
import { showDialog } from './frame';

/** Group Field: group a numeric row or column field by ranges (starting at, ending at, by). */
export function openPivotGroup(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const pivot = t ? pivotTableAt(t.ws, t.active) : undefined;
	if (!t || !pivot) return Promise.resolve(undefined);
	const fields = pivotGroupableFields(pivot);
	if (!fields.length) return Promise.resolve(undefined);
	const chosen = select(
		ctx,
		fields.map((name) => [name, name] as const),
		fields[0],
		true,
	);
	const start = numberInput(ctx, 0, -1_000_000_000, 1_000_000_000);
	const end = numberInput(ctx, 0, -1_000_000_000, 1_000_000_000);
	const by = numberInput(ctx, 1, 1, 1_000_000_000);
	const load = () => {
		const group = defaultFieldGroup(pivot, chosen.value);
		start.value = String(group.start);
		end.value = String(group.end);
		by.value = String(group.by);
	};
	chosen.addEventListener('change', load);
	load();
	return showDialog(ctx, {
		name: 'pivot-group',
		heading: 'Group Field',
		body: [
			field(ctx, 'Field:', chosen),
			field(ctx, 'Starting at:', start),
			field(ctx, 'Ending at:', end),
			field(ctx, 'By:', by),
		],
		opened: () => start.focus(),
		submit: () => {
			const group: PivotFieldGroup = {
				field: chosen.value,
				start: Number(start.value),
				end: Number(end.value),
				by: Number(by.value),
			};
			const others = (pivot.fieldGroups ?? []).filter((entry) => entry.field !== group.field);
			t.session.refreshPivotTable(t.sheet, pivot.name, { fieldGroups: [...others, group] });
			return true;
		},
	});
}
