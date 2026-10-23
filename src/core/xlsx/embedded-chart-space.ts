// The chart of an embedded workbook (an OLE chart object: Visio's Insert > Chart, or a chart another
// document embeds) read and written as a neutral `ChartSpace`, so a host edits it through
// `chart/edit` and the chart data sheet (`chart/edit/data-sheet`) and embeds the result again.
//
// Writing keeps everything else in the workbook. When the chart's data changed, the data block is
// written at A1 of the chart's sheet (series names across row 1, categories down column A), as
// Insert > Chart lays it out, and every series' references are pointed at it; a chart whose data
// lived elsewhere has that data replaced by the block. Formatting-only edits leave the cells alone.
import JSZip from 'jszip';
import { chartSpaceDataSheet } from '../chart/edit/data-sheet';
import { editChartSpace } from '../chart/edit/draft';
import { categorySource } from '../chart/edit/series-data';
import { setChartCategories, setChartSeriesValues } from '../chart/edit/series';
import type { ChartColorScheme } from '../chart/color-palettes';
import type { ChartDataSheet } from '../chart/data-sheet';
import { chartColorScheme } from '../chart/render/chart-colors';
import type { ChartSpace } from '../chart/model';
import { parseChartSpace } from '../chart/parse-space';
import { writeChartSpace } from '../chart/write-space';
import { columnLabel, quoteSheetName } from './address';
import { createEditSession } from './edit/session';
import type { CellValue, Workbook } from './model';
import { loadXlsx } from './read/index';
import { saveXlsx } from './write/index';

export interface EmbeddedChartSpace {
	/** The chart part (`xl/charts/chart1.xml`). */
	partName: string;
	/** Index of the sheet whose drawing holds the chart. */
	sheet: number;
	space: ChartSpace;
	/** The workbook theme's colours (`accent1`, `dk1`, ... as `#RRGGBB`), to resolve scheme colours. */
	scheme: ChartColorScheme;
}

const decoder = new TextDecoder();

function locate(workbook: Workbook): EmbeddedChartSpace | undefined {
	for (const [sheet, ws] of workbook.sheets.entries())
		for (const drawing of ws.drawings) {
			if (drawing.kind !== 'chart' || !drawing.partName) continue;
			const bytes = workbook.source?.parts.get(drawing.partName);
			if (!bytes) continue;
			const { chartSpace } = parseChartSpace(decoder.decode(bytes));
			if (chartSpace.plotArea.groups.length)
				return {
					partName: drawing.partName,
					sheet,
					space: chartSpace,
					scheme: chartColorScheme(workbook.theme),
				};
		}
	return undefined;
}

/** The first chart of an embedded workbook as a `ChartSpace`, or undefined when there is none. */
export async function readEmbeddedChartSpace(
	bytes: Uint8Array,
): Promise<EmbeddedChartSpace | undefined> {
	try {
		return locate(await loadXlsx(bytes));
	} catch {
		return undefined;
	}
}

const sameSheet = (a: ChartDataSheet, b: ChartDataSheet) => JSON.stringify(a) === JSON.stringify(b);

/** `space` with its data written at A1 of sheet `sheet` and every series referring to it. */
function writeDataBlock(
	workbook: Workbook,
	sheet: number,
	space: ChartSpace,
	previous: ChartDataSheet,
): ChartSpace {
	const ws = workbook.sheets[sheet]!;
	const data = chartSpaceDataSheet(space);
	const session = createEditSession(workbook);
	const rows = Math.max(previous.categories.length, data.categories.length) + 1;
	const cols = Math.max(previous.series.length, data.series.length) + 1;
	session.clearRange(
		sheet,
		{ start: { row: 0, col: 0 }, end: { row: rows - 1, col: cols - 1 } },
		'contents',
	);
	const block: CellValue[][] = [
		[null, ...data.series.map((series) => series.name)],
		...data.categories.map((label, row) => [
			label,
			...data.series.map((series) => series.values[row] ?? null),
		]),
	];
	session.setRangeValues(sheet, { row: 0, col: 0 }, block);
	const prefix = `${quoteSheetName(ws.name)}!`;
	const last = data.categories.length + 1;
	let next = setChartCategories(space, data.categories, { formula: `${prefix}$A$2:$A$${last}` });
	data.series.forEach((series, index) => {
		const col = columnLabel(index + 1);
		next = setChartSeriesValues(next, index, series.values, {
			formula: `${prefix}$${col}$2:$${col}$${last}`,
		});
	});
	return editChartSpace(next, (draft) => {
		draft.plotArea.groups
			.flatMap((group) => group.series)
			.forEach((series, index) => {
				const name = data.series[index]?.name ?? '';
				series.tx = {
					reference: categorySource([name], undefined, `${prefix}$${columnLabel(index + 1)}$1`),
					text: name,
				};
			});
	});
}

/**
 * The embedded workbook `bytes` with its first chart replaced by `space` (from
 * {@link readEmbeddedChartSpace}, edited). Throws when the workbook has no chart.
 */
export async function writeEmbeddedChartSpace(
	bytes: Uint8Array,
	space: ChartSpace,
): Promise<Uint8Array> {
	const workbook = await loadXlsx(bytes);
	const found = locate(workbook);
	if (!found) throw new Error('The embedded workbook has no chart.');
	const previous = chartSpaceDataSheet(found.space);
	const dataChanged = !sameSheet(previous, chartSpaceDataSheet(space));
	const next = dataChanged ? writeDataBlock(workbook, found.sheet, space, previous) : space;
	const zip = await JSZip.loadAsync(dataChanged ? await saveXlsx(workbook) : bytes);
	// No folder entries: an Open XML package lists parts only.
	zip.file(found.partName, writeChartSpace(next, { declarationBreak: '\n' }), {
		createFolders: false,
	});
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
