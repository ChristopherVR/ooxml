import {
	chartDrawingColorFromHex,
	setChartTargetText,
	setChartLegendEntries,
	type ChartTextEdit,
} from '../../../chart';
import type { PptxChartData, PptxChartLegendTextStyle } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same, revertStyle } from './chart-neutral-save-common';

/** Only changed resolved properties are written; untouched theme tokens stay authored. */
export function neutralTextEdit(
	before: PptxChartLegendTextStyle = {},
	after: PptxChartLegendTextStyle = {},
): ChartTextEdit {
	const edit: ChartTextEdit = {};
	for (const [host, core] of [
		['fontFamily', 'typeface'],
		['eastAsiaFontFamily', 'eastAsiaTypeface'],
		['fontSize', 'sizePt'],
	] as const)
		if (before[host] !== after[host]) Object.assign(edit, { [core]: after[host] ?? null });
	for (const key of ['bold', 'italic'] as const)
		if (before[key] !== after[key] && after[key] !== undefined) edit[key] = after[key];
	if (before.color !== after.color) {
		const color = after.color && chartDrawingColorFromHex(after.color);
		if (after.color && !color) throw new RangeError('Invalid chart text colour');
		edit.fill = color ? { kind: 'solid', color } : null;
	}
	return edit;
}

export function planNeutralChartText(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	for (const [key, target] of [
		['legendTextStyle', 'legend'],
		['textStyle', 'chartArea'],
	] as const) {
		if (same(before.style?.[key], probe.style?.[key])) continue;
		const edit = neutralTextEdit(before.style?.[key], probe.style?.[key]);
		edits.push((space) => setChartTargetText(space, { kind: target }, edit));
		revertStyle(probe, before, [key]);
	}
	const title = (data: PptxChartData): PptxChartLegendTextStyle => ({
		fontFamily: data.style?.titleFontFamily,
		eastAsiaFontFamily: data.style?.titleEastAsiaFontFamily,
		fontSize: data.style?.titleFontSize,
		bold: data.style?.titleFontBold,
		color: data.style?.titleFontColor,
	});
	if (!same(title(before), title(probe))) {
		const edit = neutralTextEdit(title(before), title(probe));
		edits.push((space) => setChartTargetText(space, { kind: 'title' }, edit));
		revertStyle(probe, before, [
			'titleFontFamily',
			'titleEastAsiaFontFamily',
			'titleFontSize',
			'titleFontBold',
			'titleFontColor',
		]);
	}
	if (!same(before.style?.legendEntries, probe.style?.legendEntries)) {
		const entries = (probe.style?.legendEntries ?? []).map((entry) => ({
			index: entry.index,
			deleted: entry.deleted,
			text: neutralTextEdit(
				before.style?.legendEntries?.find((old) => old.index === entry.index)?.textStyle,
				entry.textStyle,
			),
		}));
		edits.push((space) => setChartLegendEntries(space, entries));
		revertStyle(probe, before, ['legendEntries']);
	}
	const keys = ['fontFamily', 'eastAsiaFontFamily', 'fontSize', 'fontBold', 'fontColor'] as const;
	for (const [index, axis] of (probe.axes ?? []).entries()) {
		const old = before.axes?.[index];
		if (!old || old.axisId !== axis.axisId || keys.every((key) => old[key] === axis[key])) continue;
		const style = (item: typeof axis): PptxChartLegendTextStyle => ({
			fontFamily: item.fontFamily,
			eastAsiaFontFamily: item.eastAsiaFontFamily,
			fontSize: item.fontSize,
			bold: item.fontBold,
			color: item.fontColor,
		});
		const edit = neutralTextEdit(style(old), style(axis));
		edits.push((space) => setChartTargetText(space, { kind: 'axis', axisIndex: index }, edit));
		for (const key of keys)
			if (old[key] === undefined) delete axis[key];
			else Object.assign(axis, { [key]: old[key] });
	}
}
