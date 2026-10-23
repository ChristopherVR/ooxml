// The chart corpus and edit kinds of the xlsx chart save parity test (`chart-save-parity.test.ts`):
// every classic chart part under `xlsx/__fixtures__` (the packaged workbooks and the native Excel
// parts the JSON evidence files carry), each loaded into a workbook as a kept part, and every edit
// kind `updateChart` took before charts were edited over `ChartSpace` (chart step C2).
import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { CHART_STYLE_PRESETS } from '../../../chart/style-presets';
import type { ChartObject, Workbook } from '../../model';
import { loadXlsx } from '../../read/index';
import { createWorkbook } from '../../workbook';
import { createEditSession } from '../../edit/session';
import type { ChartPatch } from '../../edit/charts';
import { chartColorPalettePatch } from '../../edit/chart-colors';
import { chartQuickLayoutPatch, chartStylePresetPatch } from '../../edit/chart-design';
import { chartElementFillPatch, chartElementGradientPatch } from '../../edit/chart-element-fill';
import { chartSeriesFillPatch } from '../../edit/chart-series-fill';
import { chartSeriesGradientPatch } from '../../edit/chart-series-gradient';
import { chartSeriesTransparencyPatch } from '../../edit/chart-series-transparency';
import type { EditSession } from '../../edit/types';
import { saveXlsx } from '../index';

const FIXTURES = path.join(import.meta.dirname, '../../__fixtures__');

/** One chart of the corpus: how to build a workbook holding it, and where it sits. */
export interface ParityChart {
	id: string;
	load: () => Promise<Workbook>;
	sheet: number;
	index: number;
}

let host: Promise<Uint8Array> | undefined;

/** The saved workbook every evidence part is put into: data in A1:C4 and one chart. */
function hostBytes(): Promise<Uint8Array> {
	host ??= (async () => {
		const workbook = createWorkbook();
		workbook.sheets[0]!.name = 'Chart data';
		const session = createEditSession(workbook);
		session.setRangeValues(0, { row: 0, col: 0 }, [
			['Month', 'Sales', 'Cost'],
			['Jan', 10, 4],
			['Feb', 20, 8],
			['Mar', 30, 9],
		]);
		session.addChart(0, {
			anchor: { from: { row: 0, col: 6, rowOffset: 0, colOffset: 0 }, ext: { cx: 6e6, cy: 4e6 } },
			chartType: 'column',
			showLegend: true,
			series: [],
		});
		return saveXlsx(workbook);
	})();
	return host;
}

/** The package of a workbook whose first sheet holds `xml` as its kept part `chart1.xml`. */
async function withChartPartBytes(xml: string): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await hostBytes());
	zip.file('xl/charts/chart1.xml', xml);
	return zip.generateAsync({ type: 'uint8array' });
}

/** A workbook whose first sheet holds `xml` as the kept part `xl/charts/chart1.xml`. */
export async function withChartPart(xml: string): Promise<Workbook> {
	return loadXlsx(await withChartPartBytes(xml));
}

type Json = Record<string, unknown>;

/** Every classic chart part the JSON evidence files carry, by file and case. */
function jsonParts(): [string, string][] {
	const out: [string, string][] = [];
	for (const name of readdirSync(FIXTURES).filter((file) => /^excel-.*\.json$/.test(file))) {
		const data = JSON.parse(readFileSync(path.join(FIXTURES, name), 'utf8')) as Json;
		const cases = Array.isArray(data['cases']) ? (data['cases'] as Json[]) : [];
		cases.forEach((entry, i) => {
			const parts = entry['parts'] as Record<string, string> | undefined;
			const xml =
				parts?.['xl/charts/chart1.xml'] ??
				(entry['chartXml'] as string | undefined) ??
				(entry['xml'] as string | undefined);
			if (typeof xml === 'string' && xml.includes('chartSpace')) out.push([`${name}#${i}`, xml]);
		});
	}
	return out;
}

/** The corpus: packaged workbooks' charts, then the evidence files' parts. */
export async function parityCharts(): Promise<ParityChart[]> {
	const charts: ParityChart[] = [];
	for (const name of readdirSync(FIXTURES).filter((file) => file.endsWith('.xlsx'))) {
		const bytes = readFileSync(path.join(FIXTURES, name));
		let book: Workbook;
		try {
			book = await loadXlsx(bytes);
		} catch {
			continue;
		}
		book.sheets.forEach((sheet, s) =>
			sheet.drawings.forEach((drawing, index) => {
				if (drawing.kind === 'chart' && drawing.partName)
					charts.push({
						id: `${name} ${drawing.partName}`,
						load: () => loadXlsx(bytes),
						sheet: s,
						index,
					});
			}),
		);
	}
	for (const [id, xml] of jsonParts()) {
		let bytes: Promise<Uint8Array> | undefined;
		const load = async () => loadXlsx(await (bytes ??= withChartPartBytes(xml)));
		charts.push({ id, load, sheet: 0, index: 0 });
	}
	return charts;
}

const seriesEdit = (chart: ChartObject, edit: (series: ChartObject['series']) => void) => {
	if (!chart.series.length) return undefined;
	const series = structuredClone(chart.series);
	edit(series);
	return { series };
};

const bars = (chart: ChartObject) => chart.chartType === 'bar' || chart.chartType === 'column';

/** One edit kind: a patch for the chart, or `undefined` where it does not apply. */
export type ParityEdit = (
	chart: ChartObject,
	session: EditSession,
	at: { sheet: number; index: number },
) => ChartPatch | undefined | 'done';

export const PARITY_EDITS: Record<string, ParityEdit> = {
	none: () => 'done',
	title: () => ({ title: 'Parity title' }),
	titleRemove: (chart) =>
		chart.title === undefined ? undefined : ({ title: undefined } as unknown as ChartPatch),
	legendHide: (chart) => (chart.showLegend ? { showLegend: false } : undefined),
	legendBottom: () => ({ showLegend: true, legendPosition: 'b' }),
	rename: (chart) =>
		seriesEdit(chart, (series) => {
			series[0]!.name = 'Renamed series';
		}),
	seriesColor: (chart) =>
		chart.series.length ? chartSeriesFillPatch(chart, 0, { rgb: 'FF0000' }) : undefined,
	seriesNoFill: (chart) => (chart.series.length ? chartSeriesFillPatch(chart, 0, null) : undefined),
	legacyColor: (chart) =>
		seriesEdit(chart, (series) => {
			series[0]!.color = { theme: 5 };
		}),
	pointColor: (chart) =>
		seriesEdit(chart, (series) => {
			series[0]!.pointColors = {
				...series[0]!.pointColors,
				1: { kind: 'srgb', value: '00FF00', transforms: [] },
			};
		}),
	seriesGradient: (chart) =>
		chart.series.length
			? chartSeriesGradientPatch(chart, 0, { kind: 'preset', id: 1 })?.patch
			: undefined,
	seriesTransparency: (chart) =>
		chart.series.length ? chartSeriesTransparencyPatch(chart, 0, 40) : undefined,
	palette: (chart) => chartColorPalettePatch(chart, chart.colorPalette === 12 ? 14 : 12),
	stylePreset: (chart) => chartStylePresetPatch(chart, CHART_STYLE_PRESETS[1]!.id),
	quickLayout: (chart) => chartQuickLayoutPatch(chart, 5),
	spacing: (chart) => (bars(chart) ? { barGapWidth: 80, barOverlap: 10 } : undefined),
	chartAreaFill: (chart) => chartElementFillPatch(chart, 'chartArea', { rgb: 'FFEEDD' }),
	plotAreaGradient: (chart) =>
		chartElementGradientPatch(chart, 'plotArea', { kind: 'preset', id: 2 })?.patch,
	titleFill: (chart) =>
		chart.title === undefined
			? undefined
			: chartElementFillPatch(chart, 'title', { rgb: 'CCDDEE' }),
	legendFill: (chart) =>
		chart.showLegend ? chartElementFillPatch(chart, 'legend', { rgb: 'DDEEFF' }) : undefined,
	valueAxisHidden: (chart) =>
		chart.axes?.some((axis) => axis.kind === 'val')
			? {
					axes: chart.axes.map((axis) => (axis.kind === 'val' ? { ...axis, deleted: true } : axis)),
				}
			: undefined,
	insertRows: (_chart, session) => {
		session.insertRows(0, 0, 1);
		return 'done';
	},
	typeLine: (chart) =>
		chart.chartType === 'line' ? { chartType: 'column' } : { chartType: 'line' },
	groupingStacked: (chart) =>
		['bar', 'column', 'line', 'area'].includes(chart.chartType)
			? { grouping: 'stacked' }
			: undefined,
	addSeries: (chart) =>
		seriesEdit(chart, (series) => {
			series.push({ ...structuredClone(series[0]!), name: 'Added series' });
		}),
	removeSeries: (chart) =>
		chart.series.length > 1 ? { series: structuredClone(chart.series.slice(1)) } : undefined,
};

/**
 * Edit kinds whose result is allowed to differ from the DOM patch path, with why: the old path
 * regenerated the whole part from the summary (losing unmodelled formatting); the operations keep
 * the part and change only what the edit names.
 */
export const REGENERATING_EDITS = new Set([
	'typeLine',
	'groupingStacked',
	'addSeries',
	'removeSeries',
]);

/**
 * Runs `edit` on a fresh copy of `chart`: the saved chart part (or `undefined` when the edit does
 * not apply) and its source.
 */
export async function runParityEdit(
	chart: ParityChart,
	edit: ParityEdit,
): Promise<{ source: string; saved: string } | undefined> {
	const book = await chart.load();
	const object = book.sheets[chart.sheet]!.drawings[chart.index] as ChartObject;
	const partName = object.partName!;
	const source = new TextDecoder().decode(book.source!.parts.get(partName));
	const session = createEditSession(book);
	let patch: ReturnType<ParityEdit>;
	try {
		patch = edit(object, session, chart);
	} catch {
		return undefined;
	}
	if (patch === undefined) return undefined;
	if (patch !== 'done') {
		try {
			session.updateChart(chart.sheet, chart.index, patch);
		} catch {
			return undefined;
		}
	}
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	const saved = await zip.file(partName)?.async('string');
	return saved === undefined ? undefined : { source, saved };
}
