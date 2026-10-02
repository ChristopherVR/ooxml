// Fill > Series: linear, growth and date series from the first cell of each row or column of the
// selection, with a step and an optional stop value. Trend fitting is not offered.
import { dateToSerial, getCell, serialToDate, styleAt } from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { checkbox, field, radios, row, textInput } from './fields.js';
import { showDialog } from './frame.js';

export type SeriesType = 'linear' | 'growth' | 'date';
export type DateUnit = 'day' | 'weekday' | 'month' | 'year';

/** `count` values starting at `start`; stops before passing `stop`. */
export function seriesValues(
	start: number,
	count: number,
	type: SeriesType,
	step: number,
	options: { stop?: number; unit?: DateUnit; date1904?: boolean } = {},
): number[] {
	const out: number[] = [start];
	let value = start;
	for (let i = 1; i < count; i++) {
		if (type === 'linear') value = start + step * i;
		else if (type === 'growth') value = value * step;
		else value = nextDate(value, step, options.unit ?? 'day', options.date1904 ?? false);
		const stop = options.stop;
		if (stop !== undefined && (step >= 0 || type === 'growth' ? value > stop : value < stop)) break;
		out.push(Math.round(value * 1e10) / 1e10);
	}
	return out;
}

function nextDate(serial: number, step: number, unit: DateUnit, date1904: boolean): number {
	if (unit === 'day') return serial + step;
	if (unit === 'weekday') {
		let s = serial;
		let left = Math.abs(step);
		const dir = step >= 0 ? 1 : -1;
		while (left > 0) {
			s += dir;
			const day = serialToDate(s, date1904).getUTCDay();
			if (day !== 0 && day !== 6) left--;
		}
		return s;
	}
	const d = serialToDate(serial, date1904);
	const months = unit === 'month' ? step : step * 12;
	const next = new Date(
		Date.UTC(
			d.getUTCFullYear(),
			d.getUTCMonth() + months,
			d.getUTCDate(),
			d.getUTCHours(),
			d.getUTCMinutes(),
			d.getUTCSeconds(),
		),
	);
	return dateToSerial(next, date1904);
}

export function openFillSeries(ctx: EditorContext): Promise<number | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const r = t.range;
	const inRows = r.end.col - r.start.col >= r.end.row - r.start.row;
	const startCell = getCell(t.ws, r.start.row, r.start.col);
	const isDate =
		typeof startCell?.value === 'number' &&
		/[dmy]/i.test(styleAt(t.workbook, startCell.styleId).numFmt.replace(/"[^"]*"/g, ''));
	const direction = radios(
		ctx,
		'Series in',
		[
			['rows', 'Rows'],
			['columns', 'Columns'],
		],
		inRows ? 'rows' : 'columns',
	);
	const type = radios(
		ctx,
		'Type',
		[
			['linear', 'Linear'],
			['growth', 'Growth'],
			['date', 'Date'],
		],
		isDate ? 'date' : 'linear',
	);
	const unit = radios(
		ctx,
		'Date unit',
		[
			['day', 'Day'],
			['weekday', 'Weekday'],
			['month', 'Month'],
			['year', 'Year'],
		],
		'day',
	);
	const trend = checkbox(ctx, 'Trend');
	trend.input.disabled = true;
	const step = textInput(ctx, '1');
	const stop = textInput(ctx, '');
	const syncUnit = (): void => {
		for (const input of unit.inputs) input.disabled = type.get() !== 'date';
	};
	for (const input of type.inputs) input.addEventListener('change', syncUnit);
	syncUnit();
	return showDialog<number>(ctx, {
		name: 'fill-series',
		heading: 'Series',
		body: [
			row(ctx, direction.element, type.element, unit.element),
			trend.wrapper,
			row(ctx, field(ctx, 'Step value:', step), field(ctx, 'Stop value:', stop)),
		],
		opened: () => step.focus(),
		submit: () => {
			const stepValue = Number(step.value.trim());
			const stopValue = stop.value.trim() ? Number(stop.value.trim()) : undefined;
			if (
				!step.value.trim() ||
				!Number.isFinite(stepValue) ||
				(stopValue !== undefined && !Number.isFinite(stopValue))
			) {
				ctx.toast(ctx.t('Enter a number.'), 'warning');
				step.focus();
				return undefined;
			}
			const kind = type.get() as SeriesType;
			const rows = direction.get() === 'rows';
			const lanes = rows ? r.end.row - r.start.row + 1 : r.end.col - r.start.col + 1;
			const length = rows ? r.end.col - r.start.col + 1 : r.end.row - r.start.row + 1;
			let written = 0;
			t.session.batch('Series', () => {
				for (let lane = 0; lane < lanes; lane++) {
					const row0 = rows ? r.start.row + lane : r.start.row;
					const col0 = rows ? r.start.col : r.start.col + lane;
					const first = getCell(t.ws, row0, col0);
					if (typeof first?.value !== 'number') continue;
					const values = seriesValues(first.value, length, kind, stepValue, {
						...(stopValue !== undefined ? { stop: stopValue } : {}),
						unit: unit.get() as DateUnit,
						date1904: t.workbook.date1904,
					});
					const grid = rows ? [values] : values.map((v) => [v]);
					t.session.setRangeValues(t.sheet, { row: row0, col: col0 }, grid);
					const numFmt = styleAt(t.workbook, first.styleId).numFmt;
					const end = rows
						? { row: row0, col: col0 + values.length - 1 }
						: { row: row0 + values.length - 1, col: col0 };
					if (numFmt !== 'General')
						t.session.applyStyle(t.sheet, [{ start: { row: row0, col: col0 }, end }], { numFmt });
					written += values.length;
				}
			});
			return written;
		},
	});
}
