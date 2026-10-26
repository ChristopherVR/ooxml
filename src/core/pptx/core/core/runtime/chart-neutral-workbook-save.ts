// Package adapter for a classic chart's structural data edit; core owns allocation and references.
import type JSZip from 'jszip';
import {
	appendChartWorkbookSheet,
	chartSeriesEntries,
	rebaseChartData,
	writeChartSpaceOverSource,
	type ChartSpace,
} from '../../../../chart';
import type { PptxChartData } from '../../types';
import type { NeutralChartSave } from '../../utils/chart-neutral-save';

/** Stage both parts before changing the package. Unsupported workbook sources fail atomically. */
export async function prepareNeutralChartWorkbook(
	zip: JSZip,
	chart: PptxChartData,
	saved: NeutralChartSave,
	resolve: (base: string, target: string) => string,
): Promise<NeutralChartSave> {
	if (!saved.workbookModel) return saved;
	const syncIds = () =>
		chartSeriesEntries(saved.workbookModel!).forEach(({ series }, index) => {
			const target = chart.series[index];
			if (target && series.index !== undefined) target.idx = series.index;
		});
	if (!chart.externalData) {
		syncIds();
		return saved;
	}
	const data = chart.externalData;
	if (data.isExternal || !data.targetPath || !chart.chartPartPath)
		throw new Error('Structural chart edits require an accessible embedded workbook.');
	const path = resolve(chart.chartPartPath, data.targetPath);
	const bytes = await zip.file(path)?.async('uint8array');
	if (!bytes) throw new Error('The embedded chart workbook is missing.');
	let model: ChartSpace | undefined;
	const allocated = await appendChartWorkbookSheet(bytes, (name) => {
		const rebased = rebaseChartData(saved.workbookModel!, name);
		model = rebased.chartSpace;
		return rebased.rows;
	});
	const xml = writeChartSpaceOverSource(saved.xml, model!);
	zip.file(path, allocated.bytes);
	syncIds();
	return { xml, workbookWrites: [] };
}
