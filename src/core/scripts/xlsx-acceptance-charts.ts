// Acceptance workbooks whose charts were edited over their ChartSpace (chart step C2): every chart
// of a fixture formatted through the edit session (`editChart` intents and `updateChart` patches),
// and a second copy whose charts change type and grouping, so Excel opens parts the operations and
// `writeChartSpace` wrote rather than the source parts. Edits a chart cannot take are skipped.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { ChartEditIntent } from '../chart/edit/intent';
import type { ChartObject, Workbook } from '../xlsx/model';
import { loadXlsx } from '../xlsx/read/index';
import { createEditSession } from '../xlsx/edit/session';
import type { ChartPatch } from '../xlsx/edit/charts';
import type { EditSession } from '../xlsx/edit/types';
import { chartColorPalettePatch } from '../xlsx/edit/chart-colors';
import { chartQuickLayoutPatch } from '../xlsx/edit/chart-design';
import { chartElementFillPatch } from '../xlsx/edit/chart-element-fill';
import { chartSeriesFillPatch } from '../xlsx/edit/chart-series-fill';

/** The fixtures with classic charts. */
export const CHART_FIXTURES = [
	'excel-features.xlsx',
	'excel-chart-families.xlsx',
	'openpyxl-features.xlsx',
];

type Edit = (chart: ChartObject) => ChartEditIntent | { patch: ChartPatch | undefined };

const FORMAT_EDITS: Edit[] = [
	() => ({ patch: { title: 'Edited chart' } }),
	(chart) => ({ patch: chartColorPalettePatch(chart, 12) }),
	(chart) => ({ patch: chartSeriesFillPatch(chart, 0, { rgb: 'C00000' }) }),
	(chart) => ({ patch: chartElementFillPatch(chart, 'chartArea', { rgb: 'F2F2F2' }) }),
	(chart) => ({ patch: chartQuickLayoutPatch(chart, 3) }),
	() => ({ op: 'legend', edit: { position: 'b' } }),
	() => ({ op: 'dataLabels', edit: { showValue: true } }),
	() => ({ op: 'line', target: { kind: 'chartArea' }, line: { widthEmu: 19050 } }),
	() => ({
		op: 'line',
		target: { kind: 'series', seriesIndex: 0 },
		line: { widthEmu: 38100 },
		merge: true,
	}),
	(chart) => {
		const axisIndex =
			chart.chartSpace?.plotArea.axes.findIndex((axis) => axis.kind === 'val') ?? -1;
		return {
			op: 'axis',
			axisIndex,
			edit: { max: 100, majorUnit: 20, numberFormat: '0.0', minorGridlines: true },
		};
	},
	(chart) => {
		const axisIndex =
			chart.chartSpace?.plotArea.axes.findIndex((axis) => axis.kind !== 'val') ?? -1;
		return { op: 'axis', axisIndex, edit: { reverse: true, majorTickMark: 'cross' } };
	},
	() => ({ op: 'groupOptions', seriesIndex: 0, edit: { gapWidth: 80 } }),
];

const TYPE_EDITS: Edit[] = [
	(chart) => ({ patch: { chartType: chart.chartType === 'line' ? 'column' : 'line' } }),
	() => ({ patch: { grouping: 'stacked' } }),
	(chart) => ({
		patch: chart.series[0]
			? { series: [...chart.series, { ...structuredClone(chart.series[0]), name: 'Copy' }] }
			: undefined,
	}),
];

/** Applies `edits` to every chart of `workbook`, skipping those a chart cannot take. */
function editCharts(session: EditSession, workbook: Workbook, edits: Edit[]): void {
	workbook.sheets.forEach((sheet, s) =>
		sheet.drawings.forEach((drawing, index) => {
			if (drawing.kind !== 'chart') return;
			for (const edit of edits) {
				const chart = sheet.drawings[index] as ChartObject;
				try {
					const next = edit(chart);
					if ('patch' in next) {
						if (next.patch) session.updateChart(s, index, next.patch);
					} else session.editChart(s, index, next);
				} catch (error) {
					if (!(error instanceof RangeError)) throw error;
				}
			}
		}),
	);
}

/** `[file, build]` pairs: a formatted and a retyped copy of each chart fixture. */
export function chartEditBuilds(fixtures: string): [string, () => Promise<Workbook>][] {
	return CHART_FIXTURES.flatMap((name): [string, () => Promise<Workbook>][] => {
		const load = () => loadXlsx(readFileSync(join(fixtures, name)));
		const build = (edits: Edit[]) => async () => {
			const workbook = await load();
			editCharts(createEditSession(workbook), workbook, edits);
			return workbook;
		};
		return [
			[`chart-formatted-${name}`, build(FORMAT_EDITS)],
			[`chart-retyped-${name}`, build(TYPE_EDITS)],
		];
	});
}
