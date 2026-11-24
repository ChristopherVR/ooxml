import {
	chartExSummaryFromModel,
	renderChartSummarySvg,
	type ChartExFamily,
} from 'ooxml-core/chart';
import {
	createEditSession,
	createWorkbook,
	formatAddress,
	saveXlsx,
	type CellValue,
	type ChartObject,
} from 'ooxml-core/xlsx';
import { buildChartEx, INSERT_CHART_EX_TYPES } from 'ooxml-core/xlsx/ui';

const EMBED_ANCHOR = { fromCol: 4, toCol: 14, rows: 20 } as const;

export const MAP_APPROXIMATION_NOTICE =
	'Map outlines here are approximate. Office draws a new map only after it adds map data through its online service.';

/** Extract the ChartEx family from a picker type id (`cx:<family>`), or undefined for classic charts. */
export function chartExFamilyFromPickerType(type: string): ChartExFamily | undefined {
	return INSERT_CHART_EX_TYPES.find(([family]) => `cx:${family}` === type)?.[0];
}

/** Map an embedded chart's type back to the Insert Chart picker id (`cx:<family>` for ChartEx). */
export function pickerTypeFromEmbeddedChart(chart: ChartObject): string {
	switch (chart.chartType) {
		case 'waterfall':
			return 'cx:waterfall';
		case 'histogram':
			return chart.histogram?.pareto ? 'cx:pareto' : 'cx:clusteredColumn';
		case 'boxWhisker':
			return 'cx:boxWhisker';
		case 'treemap':
			return 'cx:treemap';
		case 'sunburst':
			return 'cx:sunburst';
		case 'funnel':
			return 'cx:funnel';
		case 'regionMap':
			return 'cx:regionMap';
		default:
			return chart.chartType;
	}
}

/** Build a scratch workbook with `rows` and a ChartEx object of `family`, or undefined when empty. */
function buildGridChartEx(
	rows: readonly (readonly CellValue[])[],
	family: ChartExFamily,
	title?: string,
) {
	const width = rows[0]?.length ?? 0;
	if (rows.length === 0 || width === 0) return undefined;
	const workbook = createWorkbook();
	createEditSession(workbook).setRangeValues(
		0,
		{ row: 0, col: 0 },
		rows.map((row) => [...row]),
	);
	const range = { start: { row: 0, col: 0 }, end: { row: rows.length - 1, col: width - 1 } };
	try {
		const built = buildChartEx(workbook, 0, range, family, title ? { title } : {});
		return { workbook, built };
	} catch {
		return undefined;
	}
}

/** Render a ChartEx grid preview SVG (`tile` omits title and legend for the gallery tile). */
export function gridChartExSvg(
	rows: readonly (readonly CellValue[])[],
	family: ChartExFamily,
	width: number,
	height: number,
	options: { title?: string; tile?: boolean } = {},
): string {
	const built = buildGridChartEx(rows, family, options.title)?.built;
	if (!built) return '';
	const summary = chartExSummaryFromModel(built.chartExSpace).summary;
	if (!summary) return '';
	const { title: _title, ...bare } = summary;
	return renderChartSummarySvg(options.tile ? { ...bare, showLegend: false } : summary, {
		width,
		height,
		...(family === 'regionMap' ? { mapApproximationLabel: MAP_APPROXIMATION_NOTICE } : {}),
	});
}

/** Build an embedded `.xlsx` workbook holding `rows` and one ChartEx chart of `family`. */
export async function gridChartExWorkbook(
	rows: readonly (readonly CellValue[])[],
	family: ChartExFamily,
	options: { title?: string } = {},
): Promise<Uint8Array | undefined> {
	const result = buildGridChartEx(rows, family, options.title);
	if (!result) return undefined;
	const { workbook, built } = result;
	built.anchor = {
		from: { row: 0, col: EMBED_ANCHOR.fromCol, colOffset: 0, rowOffset: 0 },
		to: { row: EMBED_ANCHOR.rows, col: EMBED_ANCHOR.toCol, colOffset: 0, rowOffset: 0 },
	};
	if (built.chartEx) built.chartEx.anchor = built.anchor;
	createEditSession(workbook).addChartEx(0, built);
	workbook.oleSize = `${formatAddress({ row: 0, col: EMBED_ANCHOR.fromCol })}:${formatAddress({
		row: EMBED_ANCHOR.rows - 1,
		col: EMBED_ANCHOR.toCol - 1,
	})}`;
	return saveXlsx(workbook);
}
