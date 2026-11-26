import {
	formatAddress,
	parseAddress,
	type PivotCondition,
	type PivotTablePatch,
} from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { selectedPivot } from '../commands/pivots';
import { assertPivotEditingAvailable } from '../commands/pivot-availability';
import { el, field, select, text, textInput } from './fields';
import { showDialog } from './frame';

export async function openPivotInput(
	ctx: EditorContext,
	kind: 'name' | 'source' | 'move',
): Promise<void> {
	const selected = selectedPivot(ctx);
	if (!selected) return;
	const { t, pivot } = selected;
	const [heading, label, original]: readonly [string, string, string] =
		kind === 'name'
			? ['PivotTable Name', 'PivotTable Name', pivot.name]
			: kind === 'source'
				? ['Change Data Source', 'Source Range:', pivot.source]
				: ['Move PivotTable', 'Destination Cell:', formatAddress(pivot.location.start)];
	const input = textInput(ctx, original);
	input.maxLength = kind === 'source' ? 512 : 255;
	await showDialog(ctx, {
		name: `pivot-${kind}`,
		heading,
		body: field(ctx, label, input),
		opened: () => {
			input.focus();
			input.select();
		},
		submit: () => {
			assertPivotEditingAvailable(ctx);
			const value = input.value.trim();
			if (value === original) return true;
			if (kind === 'source') {
				t.session.refreshPivotTable(t.sheet, pivot.name, { source: value });
				return true;
			}
			let patch: Omit<PivotTablePatch, 'source'>;
			if (kind === 'move') {
				const destination = parseAddress(value.replaceAll('$', ''));
				if (!destination) throw new Error('The reference is not valid.');
				patch = { destination };
			} else patch = { [kind]: value };
			t.session.updatePivotTable(t.sheet, pivot.name, patch);
			if (patch.destination)
				ctx.selection.set({
					active: patch.destination,
					anchor: patch.destination,
					ranges: [{ start: patch.destination, end: patch.destination }],
				});
			return true;
		},
	});
}

export async function openPivotCondition(ctx: EditorContext): Promise<void> {
	const selected = selectedPivot(ctx);
	if (!selected) return;
	const { t, pivot } = selected;
	const current = pivot.condition;
	const kind = select(
		ctx,
		[
			['', '(None)'],
			['captionEqual', 'Label Equals'],
			['valueGreaterThan', 'Value Greater Than'],
			['valueLessThan', 'Value Less Than'],
		],
		current?.kind ?? '',
	);
	const measures = [
		pivot.valueCaption,
		...(pivot.additionalValues ?? []).map(({ caption }) => caption),
	];
	const measure = select(
		ctx,
		measures.map((name, index) => [String(index), name]),
		String(current && 'measure' in current ? current.measure : 0),
		true,
	);
	const value = textInput(ctx, String(current?.value ?? ''));
	value.maxLength = 255;
	const measureField = field(ctx, 'Value Field:', measure);
	const valueField = field(ctx, 'Value:', value);
	const sync = () => {
		measureField.hidden = !kind.value.startsWith('value');
		valueField.hidden = !kind.value;
	};
	kind.addEventListener('change', sync);
	sync();
	const title = el(ctx, 'p');
	title.textContent = pivot.rowField;
	await showDialog(ctx, {
		name: 'pivot-condition',
		heading: 'Label and Value Filters',
		body: [title, field(ctx, 'Filter:', kind), measureField, valueField],
		opened: () => kind.focus(),
		submit: () => {
			assertPivotEditingAvailable(ctx);
			let condition: PivotCondition | null = null;
			if (kind.value === 'captionEqual') condition = { kind: kind.value, value: value.value };
			else if (kind.value === 'valueGreaterThan' || kind.value === 'valueLessThan') {
				const number = Number(value.value);
				if (!value.value.trim() || !Number.isFinite(number))
					throw new Error('Choose a measure and enter a finite filter value.');
				condition = { kind: kind.value, measure: Number(measure.value), value: number };
			}
			t.session.updatePivotTable(t.sheet, pivot.name, { condition });
			return true;
		},
	});
}

export async function openPivotConnections(ctx: EditorContext): Promise<void> {
	const selected = selectedPivot(ctx);
	if (!selected) return;
	const { t, pivot } = selected;
	const book = t.workbook.slicers;
	if (!book) return;
	const onPivot = (refs: readonly { name: string; tabId: number }[]) =>
		refs.some((ref) => ref.name === pivot.name && ref.tabId === t.ws.sheetId);
	const slicers = new Set(
		book.caches
			.filter(({ source }) => source.kind === 'pivot' && onPivot(source.pivots))
			.map(({ name }) => name),
	);
	const timelines = new Set(
		book.timelineCaches.filter(({ pivots }) => onPivot(pivots)).map(({ name }) => name),
	);
	const list = el(ctx, 'ul');
	for (const entry of [
		...book.slicers.filter(({ cache }) => slicers.has(cache)),
		...book.timelines.filter(({ cache }) => timelines.has(cache)),
	]) {
		const item = el(ctx, 'li');
		item.textContent = entry.caption ?? entry.name;
		list.append(item);
	}
	await showDialog(ctx, {
		name: 'pivot-connections',
		heading: 'Filter Connections',
		body: list,
		okLabel: null,
	});
}

export async function openCalculatedField(ctx: EditorContext): Promise<void> {
	const selected = selectedPivot(ctx);
	if (!selected) return;
	const { t, pivot } = selected;
	const current = pivot.calculatedField;
	const name = textInput(ctx, current?.name ?? '');
	name.maxLength = 255;
	const fields = pivot.fields.filter(
		(field) => /^[A-Za-z_][A-Za-z0-9_]{0,254}$/.test(field) && !pivot.dateFields?.includes(field),
	);
	const source = select(
		ctx,
		fields.map((field) => [field, field]),
		current?.field ?? fields[0],
		true,
	);
	const factor = textInput(ctx, String(current?.factor ?? 1));
	factor.maxLength = 32;
	await showDialog(ctx, {
		name: 'pivot-calculated-field',
		heading: 'Calculated Field',
		body: [
			field(ctx, 'Name:', name),
			field(ctx, 'Source Field:', source),
			field(ctx, 'Multiplier:', factor),
			text(
				ctx,
				current
					? 'Clear the name to remove the calculated field.'
					: 'The result is the source field total times the multiplier.',
			),
		],
		opened: () => name.focus(),
		submit: () => {
			assertPivotEditingAvailable(ctx);
			const label = name.value.trim();
			const number = Number(factor.value);
			if ((!label && !current) || (label && (!factor.value.trim() || !Number.isFinite(number))))
				throw new Error(
					'Choose a unique calculated field name, a source field and a finite multiplier.',
				);
			t.session.updatePivotTable(t.sheet, pivot.name, {
				calculatedField: label ? { name: label, field: source.value, factor: number } : null,
			});
			return true;
		},
	});
}
