// Embedded ChartEx workbooks reuse the worksheet source/cell transaction, not a host chart model.
import type { ChartExSpace } from '../chart/ex/model';
import type { ChartParseIssue } from '../chart/model-series';
import { writeChartExSpaceOverSource } from '../chart/ex/keep-source';
import { chartColorScheme } from '../chart/render/chart-colors';
import type { ChartColorScheme } from '../chart/color-palettes';
import type { Workbook } from './model';
import { loadXlsx } from './read';
import { saveXlsx } from './write';
import { createEditSession } from './edit/session';
import { xlsxChartExSpace } from './edit/chart-ex';
export interface EmbeddedChartExSpace {
	partName: string;
	sheet: number;
	drawing: number;
	space: ChartExSpace;
	sourceXml: string;
	issues: ChartParseIssue[];
	editable: boolean;
	scheme: ChartColorScheme;
}
function locate(book: Workbook): EmbeddedChartExSpace | undefined {
	for (const [sheet, ws] of book.sheets.entries())
		for (const [index, drawing] of ws.drawings.entries()) {
			if (drawing.kind !== 'unsupported' || !drawing.chartExSpace || !drawing.chartExPartName)
				continue;
			const bytes = book.source?.parts.get(drawing.chartExPartName);
			if (!bytes) continue;
			let space = drawing.chartExSpace;
			try {
				space = xlsxChartExSpace(book, sheet, index);
			} catch {
				/* Unavailable references retain their caches. */
			}
			return {
				partName: drawing.chartExPartName,
				sheet,
				drawing: index,
				space,
				sourceXml: new TextDecoder().decode(bytes),
				issues: drawing.chartExIssues ?? [],
				editable: !drawing.chartExIssues?.length,
				scheme: chartColorScheme(book.theme),
			};
		}
	return undefined;
}
/** The first extended chart of an embedded workbook, with live local source data when available. */
export async function readEmbeddedChartExSpace(
	bytes: Uint8Array,
): Promise<EmbeddedChartExSpace | undefined> {
	try {
		return locate(await loadXlsx(bytes));
	} catch {
		return undefined;
	}
}
/** Save through the existing worksheet ChartEx transaction, retaining untouched package parts. */
export async function writeEmbeddedChartExSpace(
	bytes: Uint8Array,
	space: ChartExSpace,
): Promise<Uint8Array> {
	const book = await loadXlsx(bytes);
	const found = locate(book);
	if (!found) throw new Error('The embedded workbook has no extended chart.');
	if (JSON.stringify(found.space) === JSON.stringify(space)) return bytes;
	// Validate before any source-cell mutation, including callers bypassing a host UI.
	writeChartExSpaceOverSource(found.sourceXml, space);
	createEditSession(book).editChartEx(found.sheet, found.drawing, () => space);
	return saveXlsx(book);
}
