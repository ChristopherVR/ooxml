import JSZip from 'jszip';
import { createWorkbook, createEditSession, saveXlsx } from 'ooxml-core/xlsx';

/** Reuse one workbook shell for native chart-part browser references. */
export async function nativeChartFixture(
	parts: Readonly<Record<string, string>>,
	minorFont: string,
	majorFont = 'Aptos Display',
): Promise<Buffer> {
	const book = createWorkbook();
	book.theme.minorFont = minorFont;
	book.theme.majorFont = majorFont;
	createEditSession(book).addChart(0, {
		chartType: 'column',
		showLegend: true,
		series: [],
		anchor: {
			from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
			ext: { cx: 600 * 9525, cy: 400 * 9525 },
		},
	});
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	for (const [part, xml] of Object.entries(parts)) zip.file(part, xml);
	return Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
}
