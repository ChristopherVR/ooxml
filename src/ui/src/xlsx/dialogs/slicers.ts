import {
	activeSlicer,
	activeTimeline,
	slicerSourceAt,
	target,
	timelineSourceAt,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
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

/** Insert Timelines: one check box per date field of the PivotTable. */
export function openInsertTimeline(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const source = timelineSourceAt(ctx);
	if (!t || !source) return Promise.resolve(undefined);
	const list = el(ctx, 'div', 'xve-checklist');
	list.setAttribute('role', 'group');
	list.setAttribute('aria-label', ctx.t('Fields'));
	const boxes = source.fields.map((name) => {
		const box = checkbox(ctx, name);
		box.wrapper.querySelector('span')!.textContent = name;
		box.input.setAttribute('aria-label', name);
		list.append(box.wrapper);
		return { name, input: box.input };
	});
	return showDialog(ctx, {
		name: 'insert-timeline',
		heading: 'Insert Timelines',
		body: [text(ctx, 'Choose the date fields you want to filter with timelines.'), list],
		opened: () => boxes[0]?.input.focus(),
		submit: () => {
			const chosen = boxes.filter((box) => box.input.checked);
			if (!chosen.length) throw new Error('Choose at least one field.');
			let last: string | undefined;
			for (const box of chosen)
				last = t.session.insertTimeline(t.sheet, { pivot: source.pivot.name, field: box.name });
			const index = t.ws.drawings.findIndex((d) => d.kind === 'timeline' && d.name === last);
			if (index >= 0) ctx.selection.set({ drawing: index });
			return true;
		},
	});
}

/** Timeline Settings: caption, visible parts and time level. */
export function openTimelineSettings(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const found = activeTimeline(ctx);
	if (!t || !found) return Promise.resolve(undefined);
	const { timeline } = found;
	const caption = textInput(ctx, timeline.caption ?? timeline.name);
	const header = checkbox(ctx, 'Display header', timeline.showHeader !== false);
	const label = checkbox(ctx, 'Show selection label', timeline.showSelectionLabel !== false);
	const levels = checkbox(ctx, 'Show time level', timeline.showTimeLevel !== false);
	const scroll = checkbox(
		ctx,
		'Show horizontal scrollbar',
		timeline.showHorizontalScrollbar !== false,
	);
	const level = select(
		ctx,
		[
			['0', 'Years'],
			['1', 'Quarters'],
			['2', 'Months'],
			['3', 'Days'],
		],
		String(timeline.level),
	);
	return showDialog(ctx, {
		name: 'timeline-settings',
		heading: 'Timeline Settings',
		body: [
			field(ctx, 'Caption:', caption),
			header.wrapper,
			label.wrapper,
			levels.wrapper,
			scroll.wrapper,
			field(ctx, 'Time Level', level),
		],
		opened: () => caption.focus(),
		submit: () => {
			t.session.updateTimeline(timeline.name, {
				caption: caption.value,
				showHeader: header.input.checked,
				showSelectionLabel: label.input.checked,
				showTimeLevel: levels.input.checked,
				showHorizontalScrollbar: scroll.input.checked,
				level: Number(level.value),
			});
			return true;
		},
	});
}
