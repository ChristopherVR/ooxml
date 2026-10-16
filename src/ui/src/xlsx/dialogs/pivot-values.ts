import type { PivotValueField } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { field, fieldset, select, textInput } from './fields';
import { button } from './frame';

export const PIVOT_AGGREGATE_OPTIONS = [
	['sum', 'Sum'],
	['count', 'Count'],
	['average', 'Average'],
	['min', 'Min'],
	['max', 'Max'],
	['product', 'Product'],
	['countNums', 'Count Numbers'],
	['stdDev', 'StdDev'],
	['stdDevp', 'StdDevp'],
	['var', 'Var'],
	['varp', 'Varp'],
] as const;

/** Extra value controls share source fields, preserve imported captions, and keep a stable order. */
export function pivotValueControls(
	ctx: EditorContext,
	initial: readonly PivotValueField[],
	changed: () => void,
	initialFields: readonly string[] = [],
) {
	const doc = ctx.host.ownerDocument;
	const host = doc.createElement('div');
	const controls: {
		element: HTMLElement;
		field: HTMLSelectElement;
		aggregate: HTMLSelectElement;
		caption: HTMLInputElement;
	}[] = [];
	let fields: string[] = [...initialFields];
	let ready = false;
	const add = button(ctx, 'Add Value Field');
	const load = (control: HTMLSelectElement, preferred = control.value) => {
		control.replaceChildren(
			...fields.map((name) => {
				const option = doc.createElement('option');
				option.value = name;
				option.textContent = name;
				return option;
			}),
		);
		control.value = fields.includes(preferred) ? preferred : (fields.at(-1) ?? '');
	};
	const append = (initial?: PivotValueField) => {
		if (controls.length >= 15) return;
		const source = select(ctx, [], initial?.field ?? '');
		load(source, initial?.field ?? fields.at(-1));
		const aggregate = select(ctx, PIVOT_AGGREGATE_OPTIONS, initial?.aggregate ?? 'sum');
		const caption = textInput(ctx, initial?.caption ?? '');
		let automaticCaption =
			!initial ||
			initial.caption ===
				`${initial.aggregate[0]!.toUpperCase()}${initial.aggregate.slice(1)} of ${initial.field}`;
		caption.addEventListener('input', () => {
			automaticCaption = false;
		});
		const defaultCaption = () => {
			caption.value = `${aggregate.value[0]!.toUpperCase()}${aggregate.value.slice(1)} of ${source.value}`;
		};
		if (!initial) defaultCaption();
		source.addEventListener('change', () => {
			if (automaticCaption) defaultCaption();
		});
		aggregate.addEventListener('change', () => {
			if (automaticCaption) defaultCaption();
		});
		const remove = button(ctx, 'Remove Value Field');
		const element = fieldset(
			ctx,
			'Additional Value Field',
			field(ctx, 'Value Field:', source),
			field(ctx, 'Summarize Values By:', aggregate),
			field(ctx, 'Value Caption:', caption),
			remove,
		);
		element.dataset.pivotValue = '';
		const row = { element, field: source, aggregate, caption };
		remove.addEventListener('click', () => {
			controls.splice(controls.indexOf(row), 1);
			element.remove();
			add.disabled = false;
			changed();
		});
		controls.push(row);
		host.insertBefore(element, add);
		add.disabled = controls.length >= 15;
		if (ready) changed();
	};
	host.append(add);
	for (const value of initial) append(value);
	ready = true;
	add.addEventListener('click', () => append());
	return {
		host,
		count: () => controls.length,
		load: (names: string[]) => {
			fields = names;
			for (const row of controls) {
				const old = row.field.value;
				const prefix = `${row.aggregate.value[0]!.toUpperCase()}${row.aggregate.value.slice(1)} of `;
				load(row.field);
				if (row.caption.value === prefix + old) row.caption.value = prefix + row.field.value;
			}
		},
		read: (): PivotValueField[] =>
			controls.map((row) => ({
				field: row.field.value,
				aggregate: row.aggregate.value as PivotValueField['aggregate'],
				caption: row.caption.value,
			})),
	};
}
