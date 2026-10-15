import { activeSlicer, slicerSourceAt, target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, el, field, select, text, textInput } from './fields';
import { showDialog } from './frame';

/** Insert Slicers: one check box per table column or PivotTable field. */
export function openInsertSlicer(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const source = slicerSourceAt(ctx);
	if (!t || !source) return Promise.resolve(undefined);
	const list = el(ctx, 'div', 'xve-checklist');
	list.setAttribute('role', 'group');
	list.setAttribute('aria-label', ctx.t('Fields'));
	const boxes = source.fields.map((name) => {
		const box = checkbox(ctx, name);
		// A field name is data, not a UI string: show it as written.
		box.wrapper.querySelector('span')!.textContent = name;
		box.input.setAttribute('aria-label', name);
		list.append(box.wrapper);
		return { name, input: box.input };
	});
	return showDialog(ctx, {
		name: 'insert-slicer',
		heading: 'Insert Slicers',
		body: [text(ctx, 'Choose the fields you want to filter with slicers.'), list],
		opened: () => boxes[0]?.input.focus(),
		submit: () => {
			const chosen = boxes.filter((box) => box.input.checked);
			if (!chosen.length) throw new Error('Choose at least one field.');
			let last: string | undefined;
			for (const box of chosen)
				last = t.session.insertSlicer(
					t.sheet,
					source.kind === 'table'
						? { kind: 'table', table: source.table.name, column: box.name }
						: { kind: 'pivot', pivot: source.pivot.name, field: box.name },
				);
			const drawings = t.ws.drawings;
			const index = drawings.findIndex((d) => d.kind === 'slicer' && d.name === last);
			if (index >= 0) ctx.selection.set({ drawing: index });
			return true;
		},
	});
}

/** Slicer Settings: caption, header, sort order and items without data. */
export function openSlicerSettings(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const found = activeSlicer(ctx);
	if (!t || !found) return Promise.resolve(undefined);
	const { slicer, cache } = found;
	const name = textInput(ctx, slicer.name);
	name.disabled = true;
	const caption = textInput(ctx, slicer.caption ?? slicer.name);
	const header = checkbox(ctx, 'Display header', slicer.showCaption !== false);
	const order = select(
		ctx,
		[
			['ascending', 'Ascending (A to Z)'],
			['descending', 'Descending (Z to A)'],
		],
		cache.sortOrder ?? 'ascending',
	);
	const hide = checkbox(ctx, 'Hide items with no data', cache.hideNoData === true);
	return showDialog(ctx, {
		name: 'slicer-settings',
		heading: 'Slicer Settings',
		body: [
			field(ctx, 'Name to use in formulas:', name),
			field(ctx, 'Caption:', caption),
			header.wrapper,
			field(ctx, 'Sort items:', order),
			hide.wrapper,
		],
		opened: () => caption.focus(),
		submit: () => {
			t.session.updateSlicer(slicer.name, {
				caption: caption.value,
				showCaption: header.input.checked,
				sortOrder: order.value as 'ascending' | 'descending',
				hideNoData: hide.input.checked,
			});
			return true;
		},
	});
}
