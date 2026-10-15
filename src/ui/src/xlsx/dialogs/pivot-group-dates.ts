import {
	PIVOT_DATE_UNITS,
	dateGroupDays,
	defaultDateGroup,
	pivotGroupableDateFields,
	pivotIsoToSerial,
	pivotSerialToIso,
	pivotTableAt,
	type PivotDateGroup,
	type PivotDateUnit,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, fieldset, numberInput, row, select, textInput } from './fields';
import { showDialog } from './frame';

const UNIT_LABELS: Record<PivotDateUnit, string> = {
	seconds: 'Seconds',
	minutes: 'Minutes',
	hours: 'Hours',
	days: 'Days',
	months: 'Months',
	quarters: 'Quarters',
	years: 'Years',
};

/** A serial as the dialog types it: the date, plus the time of day when it has one. */
const shown = (serial: number, date1904: boolean): string =>
	pivotSerialToIso(serial, date1904)
		.replace(/T00:00:00$/, '')
		.replace('T', ' ');

/**
 * Grouping (Group Selection on a date field): group a date field by seconds, minutes, hours, days,
 * months, quarters and years, from an automatic or chosen start to an automatic or chosen end.
 */
export function openPivotGroupDates(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const pivot = t ? pivotTableAt(t.ws, t.active) : undefined;
	const workbook = ctx.workbook();
	if (!t || !pivot || !workbook) return Promise.resolve(undefined);
	const date1904 = workbook.date1904;
	const groupable = pivotGroupableDateFields(workbook, t.sheet, pivot);
	const fields = [
		...new Set([...(pivot.dateGroups ?? []).map((group) => group.field), ...groupable]),
	].filter((name) => pivot.fields.includes(name));
	if (!fields.length) return Promise.resolve(undefined);
	const chosen = select(
		ctx,
		fields.map((name) => [name, name] as const),
		fields[0],
		true,
	);
	const autoStart = checkbox(ctx, 'Auto', true);
	const autoEnd = checkbox(ctx, 'Auto', true);
	const start = textInput(ctx);
	const end = textInput(ctx);
	const days = numberInput(ctx, 1, 1, 32_767);
	const units = PIVOT_DATE_UNITS.map((unit) => ({
		unit,
		box: checkbox(ctx, UNIT_LABELS[unit], false),
	}));
	const selected = () =>
		units.filter((entry) => entry.box.input.checked).map((entry) => entry.unit);
	const refresh = () => {
		start.disabled = autoStart.input.checked;
		end.disabled = autoEnd.input.checked;
		// Several units count single days; only a lone Days choice takes a length.
		days.disabled = selected().join() !== 'days';
		if (days.disabled) days.value = '1';
	};
	const load = () => {
		const group = defaultDateGroup(pivot, chosen.value);
		const existing = pivot.dateGroups?.find((entry) => entry.field === chosen.value);
		autoStart.input.checked = existing?.start === undefined;
		autoEnd.input.checked = existing?.end === undefined;
		start.value = shown(group.start, date1904);
		// An automatic range of whole days ends a day after the latest date: show that date.
		end.value = shown(
			existing?.end === undefined && Number.isInteger(group.end) ? group.end - 1 : group.end,
			date1904,
		);
		for (const entry of units)
			entry.box.input.checked = (existing?.by ?? ['months']).includes(entry.unit);
		days.value = String(existing ? dateGroupDays(existing) : 1);
		refresh();
	};
	chosen.addEventListener('change', load);
	for (const input of [autoStart.input, autoEnd.input, ...units.map((entry) => entry.box.input)])
		input.addEventListener('change', refresh);
	load();
	const parse = (text: string): number => {
		const serial = pivotIsoToSerial(text.trim().replace(' ', 'T'), date1904);
		if (serial === undefined) throw new Error('Enter a valid date.');
		return serial;
	};
	return showDialog(ctx, {
		name: 'pivot-group-dates',
		heading: 'Grouping',
		body: [
			field(ctx, 'Field:', chosen),
			row(ctx, field(ctx, 'Starting at:', start), autoStart.wrapper),
			row(ctx, field(ctx, 'Ending at:', end), autoEnd.wrapper),
			fieldset(ctx, 'By:', ...units.map((entry) => entry.box.wrapper)),
			field(ctx, 'Number of days:', days),
		],
		opened: () => chosen.focus(),
		submit: () => {
			const by = selected();
			if (!by.length) throw new Error('Choose at least one group.');
			const group: PivotDateGroup = {
				field: chosen.value,
				by,
				...(autoStart.input.checked ? {} : { start: parse(start.value) }),
				...(autoEnd.input.checked ? {} : { end: parse(end.value) }),
				...(by.join() === 'days' && Number(days.value) > 1 ? { days: Number(days.value) } : {}),
			};
			const others = (pivot.dateGroups ?? []).filter((entry) => entry.field !== group.field);
			t.session.refreshPivotTable(t.sheet, pivot.name, { dateGroups: [...others, group] });
			return true;
		},
	});
}
