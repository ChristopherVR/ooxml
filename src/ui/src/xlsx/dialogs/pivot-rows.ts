import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { field, fieldset, select } from './fields';
import { button } from './frame';

const MAX_EXTRA_ROWS = 62;

const LABELS = {
	row: {
		add: 'Add Row Field',
		remove: 'Remove Row Field',
		group: 'Additional Row Field',
		field: 'Row Field:',
		marker: 'pivotRow',
	},
	column: {
		add: 'Add Column Field',
		remove: 'Remove Column Field',
		group: 'Additional Column Field',
		field: 'Column Field:',
		marker: 'pivotColumn',
	},
} as const;

/** The third and following row fields (see `pivotFieldListControls`). */
export const pivotRowControls = (
	ctx: EditorContext,
	initial: readonly string[],
	changed: () => void,
	initialFields: readonly string[] = [],
) => pivotFieldListControls(ctx, 'row', initial, changed, initialFields);

/** The second and following column fields (see `pivotFieldListControls`). */
export const pivotColumnControls = (
	ctx: EditorContext,
	initial: readonly string[],
	changed: () => void,
	initialFields: readonly string[] = [],
) => pivotFieldListControls(ctx, 'column', initial, changed, initialFields);

/**
 * Further fields of one axis: one chooser per field, in nesting order, each with its own remove
 * button. The choosers share the source fields of the first ones.
 */
function pivotFieldListControls(
	ctx: EditorContext,
	axis: keyof typeof LABELS,
	initial: readonly string[],
	changed: () => void,
	initialFields: readonly string[] = [],
) {
	const text = LABELS[axis];
	const doc = ctx.host.ownerDocument;
	const host = doc.createElement('div');
	const controls: { element: HTMLElement; field: HTMLSelectElement }[] = [];
	let fields: string[] = [...initialFields];
	let ready = false;
	const add = button(ctx, text.add);
	const load = (control: HTMLSelectElement, preferred = control.value) => {
		control.replaceChildren(
			...fields.map((name) => {
				const option = doc.createElement('option');
				option.value = name;
				option.textContent = name;
				return option;
			}),
		);
		control.value = fields.includes(preferred) ? preferred : (fields[0] ?? '');
	};
	const append = (name?: string) => {
		if (controls.length >= MAX_EXTRA_ROWS) return;
		const source = select(ctx, [], name ?? '');
		load(source, name ?? fields.find((candidate) => !used().includes(candidate)));
		const remove = button(ctx, text.remove);
		const element = fieldset(ctx, text.group, field(ctx, text.field, source), remove);
		element.dataset[text.marker] = '';
		const entry = { element, field: source };
		source.addEventListener('change', () => changed());
		remove.addEventListener('click', () => {
			controls.splice(controls.indexOf(entry), 1);
			element.remove();
			add.disabled = false;
			changed();
		});
		controls.push(entry);
		host.insertBefore(element, add);
		add.disabled = controls.length >= MAX_EXTRA_ROWS;
		if (ready) changed();
	};
	const used = () => controls.map((entry) => entry.field.value);
	host.append(add);
	for (const name of initial) append(name);
	ready = true;
	add.addEventListener('click', () => append());
	return {
		host,
		count: () => controls.length,
		/** Offers these source fields to every chooser, keeping the ones still available. */
		load: (names: string[]) => {
			fields = names;
			for (const entry of controls) load(entry.field);
		},
		/** Hides the whole group (and keeps its choices) while the second row field is empty. */
		enable: (enabled: boolean) => {
			add.disabled = !enabled || controls.length >= MAX_EXTRA_ROWS;
			for (const entry of controls) entry.field.disabled = !enabled;
		},
		read: (): string[] => controls.map((entry) => entry.field.value),
	};
}
