// Package adapter for a classic chart's structural data edit; core owns allocation and references.
import type JSZip from 'jszip';
import {
	appendChartWorkbookSheet,
	chartIssuesLostOnWrite,
	chartSeriesEntries,
	chartSpaceRangeConflicts,
	parseChartSpace,
	rebaseChartData,
	writeChartSpaceOverSource,
	type ChartSpace,
} from '../../../../chart';
import type { PptxChartData } from '../../types';
import type { NeutralChartSave } from '../../utils/chart-neutral-save';

type ResolvePath = (base: string, target: string) => string;

/** Stage both parts before changing the package. Unsupported workbook sources fail atomically. */
async function prepareWorkbook(
	zip: JSZip,
	chart: PptxChartData,
	saved: NeutralChartSave,
	resolve: ResolvePath,
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
	return { xml, workbookRebased: true };
}

export class NeutralChartSaveRefusedError extends Error {}

export async function prepareNeutralChartWorkbook(
	...args: Parameters<typeof prepareWorkbook>
): Promise<NeutralChartSave> {
	try {
		return await prepareWorkbook(...args);
	} catch (error) {
		throw new NeutralChartSaveRefusedError(error instanceof Error ? error.message : String(error));
	}
}

/**
 * `xml`, a chart part about to be saved, with its data moved onto a fresh sheet of its embedded
 * workbook when its ranges cannot hold its caches where those of the `earlier` parts (the one the
 * save started from, the one the workbook is in step with) could (`chartSpaceRangeConflicts`): a
 * writer that changed the rows or series without new ranges, or data saved over a part it was not
 * edited from. `undefined` when the ranges hold the caches, or when the part has no embedded
 * workbook or cannot be written by the shared writer; the cell-for-cell write-back then reports
 * what it cannot write.
 */
export async function rebaseConflictingChartRanges(
	zip: JSZip,
	chart: PptxChartData,
	xml: string,
	earlier: readonly (string | undefined)[],
	resolve: ResolvePath,
): Promise<string | undefined> {
	if (!chart.externalData || chart.externalData.isExternal) return undefined;
	try {
		const { chartSpace, issues } = parseChartSpace(xml);
		if (chartIssuesLostOnWrite(issues).length > 0) return undefined;
		const before = earlier.flatMap((part) =>
			part === undefined ? [] : [parseChartSpace(part).chartSpace],
		);
		if (chartSpaceRangeConflicts(chartSpace, ...before).length === 0) return undefined;
		return (await prepareWorkbook(zip, chart, { xml, workbookModel: chartSpace }, resolve)).xml;
	} catch {
		return undefined;
	}
}
