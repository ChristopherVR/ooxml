// Charts in the editor: what Insert Chart and Change Chart Type offer, and the edits Chart Design
// makes to the selected chart. The chart model and its operations are the shared core's
// (`ooxml-core/chart`, `ooxml-core/docx`); this module maps picker choices onto them and
// dispatches the result as one undo step (`replaceSelectedChart`).
import type { EditorView } from 'prosemirror-view';
import {
	setChartExLegend,
	setChartLegend,
	setChartType,
	type ChartTypeTarget,
} from 'ooxml-core/chart';
import { editDocxChartEx, editDocxChart, type DocxChart } from 'ooxml-core/docx';
import { replaceSelectedChart, selectedChart } from 'ooxml-core/docx/ui';
import type { LocalizationKey } from './localization';

/** The chart types the picker offers, with their label keys, in Word's gallery order. */
export const DOCX_CHART_TYPES: ReadonlyArray<readonly [id: string, label: LocalizationKey]> = [
	['column', 'chart.type.column'],
	['bar', 'chart.type.bar'],
	['line', 'chart.type.line'],
	['area', 'chart.type.area'],
	['pie', 'chart.type.pie'],
	['doughnut', 'chart.type.doughnut'],
	['scatter', 'chart.type.scatter'],
	['radar', 'chart.type.radar'],
];

export const DOCX_CHART_GROUPINGS: ReadonlyArray<readonly [id: string, label: LocalizationKey]> = [
	['clustered', 'Clustered'],
	['stacked', 'Stacked'],
	['percentStacked', '100% Stacked'],
];

/** Types whose grouping the picker offers. */
export const DOCX_GROUPED_CHART_TYPES = ['column', 'bar'];

/** The type switch a picker choice asks for. */
export function chartTypeTarget(type: string, grouping: string): ChartTypeTarget {
	const bars = type === 'column' || type === 'bar';
	if (bars)
		return {
			kind: 'bar',
			barDirection: type === 'bar' ? 'bar' : 'col',
			grouping: grouping === 'stacked' || grouping === 'percentStacked' ? grouping : 'clustered',
		};
	const kind = DOCX_CHART_TYPES.some(([id]) => id === type) ? type : 'bar';
	return { kind: kind as ChartTypeTarget['kind'] };
}

/** The picker choice that describes `chart` (Change Chart Type opens on it). */
export function chartPickerChoice(chart: DocxChart): { type: string; grouping: string } {
	const group = chart.chartSpace?.plotArea.groups[0];
	const kind = group?.kind ?? 'bar';
	const type = kind === 'bar' ? (group?.barDirection === 'bar' ? 'bar' : 'column') : String(kind);
	const grouping =
		group?.grouping === 'stacked' || group?.grouping === 'percentStacked'
			? group.grouping
			: 'clustered';
	return { type, grouping };
}

/** Applies `edit` to the selected chart; false when no editable chart is selected. */
export function editSelectedChart(
	view: EditorView,
	edit: (chart: DocxChart) => DocxChart,
): boolean {
	const selected = selectedChart(view.state);
	if (!view.editable || !selected?.editable) return false;
	const next = edit(selected.chart);
	if (next === selected.chart) return true;
	const tr = replaceSelectedChart(view.state, next);
	if (!tr) return false;
	view.dispatch(tr);
	return true;
}

/** Change Chart Type for the selected chart. */
export function setSelectedChartType(view: EditorView, type: string, grouping: string): boolean {
	return editSelectedChart(view, (chart) =>
		editDocxChart(chart, (space) => setChartType(space, chartTypeTarget(type, grouping))),
	);
}

/** Chart Design > Legend: `none`, or a position (`r`, `t`, `l`, `b`). */
export function setSelectedChartLegend(view: EditorView, value: string): boolean {
	return editSelectedChart(view, (chart) =>
		chart.chartExSpace
			? editDocxChartEx(chart, (space) =>
					setChartExLegend(space, { visible: value !== 'none', position: value }),
				)
			: editDocxChart(chart, (space) =>
					value === 'none'
						? setChartLegend(space, { visible: false })
						: setChartLegend(space, { position: value as 'r' | 't' | 'l' | 'b' }),
				),
	);
}
