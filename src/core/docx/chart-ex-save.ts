// Reconcile edited ChartEx caches with the embedded workbook atomically before publishing its part.
import type JSZip from 'jszip';
import {
	chartExSpaceWorkbookWrites,
	chartExDataChanged,
	parseChartExSpace,
	saveChartWorkbookUpdates,
	writeChartExSpaceOverSource,
} from '../chart';
import { chartsIn, documentDrawingBlocks } from './chart';
import type { DocumentModel } from './model';

export async function applyChartExPartEdits(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel | undefined,
): Promise<void> {
	if (!base) return;
	const loaded = new Map(
		chartsIn(documentDrawingBlocks(base))
			.filter((c) => c.partName && c.chartExSpace)
			.map((c) => [c.partName!, JSON.stringify(c.chartExSpace)]),
	);
	const saved = new Set<string>();
	for (const chart of chartsIn(documentDrawingBlocks(model))) {
		if (!chart.partName || !chart.chartExSpace || saved.has(chart.partName)) continue;
		const original = loaded.get(chart.partName);
		if (original === undefined || original === JSON.stringify(chart.chartExSpace)) continue;
		const source = await zip.file(chart.partName)?.async('string');
		if (!source) throw new Error('The source chart part is missing.');
		const written = writeChartExSpaceOverSource(source, chart.chartExSpace);
		const before = parseChartExSpace(source).chartExSpace;
		if (chartExDataChanged(before, chart.chartExSpace)) {
			if (chart.chartExSpace.chartData.externalData && !chart.workbookPartName)
				throw new Error('Extended chart data cannot update an external or missing workbook.');
			const writes = chartExSpaceWorkbookWrites(chart.chartExSpace);
			await saveChartWorkbookUpdates(
				{
					read: async (part) => zip.file(part)?.async('uint8array'),
					write: (part, bytes) => {
						zip.file(part, bytes);
					},
				},
				chart.partName,
				chart.chartExSpace.chartData.externalData
					? chart.workbookPartName
						? { partName: chart.workbookPartName }
						: { external: true }
					: undefined,
				writes,
				{ atomic: true },
			);
		}
		zip.file(chart.partName, written);
		saved.add(chart.partName);
	}
}
