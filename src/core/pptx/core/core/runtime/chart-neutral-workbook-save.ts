// Package adapter for a classic chart's structural data edit; core owns allocation and references.
import type JSZip from 'jszip';
import {
	appendChartWorkbookSheet,
	chartIssuesLostOnWrite,
	chartSeriesEntries,
	chartSpaceRangeConflicts,
	openEmbeddedPackage,
	rebaseChartData,
	worksheetPathsByName,
	type ChartSpace,
} from '../../../../chart';
import { readChartSpaceSource, writeParsedChartSpace } from '../../../../chart/keep-source';
import { chartSpaceWorkbookWrites } from '../../../../chart/workbook/space-writes';
import type { PptxChartData, XmlObject } from '../../types';
import type { NeutralChartSave } from '../../utils/chart-neutral-save';
import {
	chartPartSheets,
	placeChartData,
	type ChartDataSheets,
} from '../../utils/chart-workbook-sheets';

type ResolvePath = (base: string, target: string) => string;
export type ReadChartSource = typeof readChartSpaceSource;
export interface ChartRefreshSource {
	readSource: ReadChartSource;
	readTree?: () => XmlObject;
}

/** A chart save can see the same source as the edit baseline and again at workbook write-back. */
export function chartSourcesForSave(): ReadChartSource {
	const sources = new Map<string, ReturnType<ReadChartSource>>();
	return (text) => {
		const previous = sources.get(text);
		if (previous) return previous;
		const source = readChartSpaceSource(text);
		sources.set(text, source);
		return source;
	};
}

export function workbookWritesForChart(readSource: ReadChartSource, before: string, after: string) {
	try {
		return chartSpaceWorkbookWrites(readSource(after).chartSpace, readSource(before).chartSpace);
	} catch {
		return [];
	}
}

/**
 * Stage both parts before changing the package. Unsupported workbook sources fail atomically.
 * The data goes on a sheet of its own (`placeChartData`).
 */
async function prepareWorkbook(
	zip: JSZip,
	chart: PptxChartData,
	saved: NeutralChartSave,
	resolve: ResolvePath,
	sheets: ChartDataSheets,
	readSource: ReadChartSource = readChartSpaceSource,
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
	const place = await placeChartData(zip, path, chart.chartPartPath, sheets);
	let model: ChartSpace | undefined;
	const written = await appendChartWorkbookSheet(
		bytes,
		(name) => {
			const rebased = rebaseChartData(saved.workbookModel!, name);
			model = rebased.chartSpace;
			return rebased.rows;
		},
		place.options,
	);
	const xml = writeParsedChartSpace(readSource(saved.xml), model!);
	zip.file(path, written.bytes);
	place.written(written.sheetName);
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
 * Whether `xml` names a sheet its embedded workbook lacks that `inStep`, the part the workbook is
 * in step with, does not: a sheet this session removed. A sheet the part named as loaded stays the
 * cell-for-cell write-back's to report.
 */
async function namesRemovedSheet(
	zip: JSZip,
	chart: PptxChartData,
	xml: string,
	inStep: string | undefined,
	resolve: ResolvePath,
): Promise<boolean> {
	const target = chart.externalData?.targetPath;
	if (!target || !chart.chartPartPath) return false;
	const bytes = await zip.file(resolve(chart.chartPartPath, target))?.async('uint8array');
	if (!bytes) return false;
	const sheets = await worksheetPathsByName(await openEmbeddedPackage(bytes));
	const named = chartPartSheets(inStep ?? '');
	return [...chartPartSheets(xml)].some((sheet) => !sheets.has(sheet) && !named.has(sheet));
}

/**
 * `xml`, a chart part about to be saved, with its data moved onto a sheet of its own in its
 * embedded workbook when its ranges cannot hold its caches where those of `from` (the part the save
 * started from) and `inStep` (the one the workbook is in step with) could
 * (`chartSpaceRangeConflicts`), or name a sheet this session removed: a writer that changed the
 * rows or series without new ranges, or data saved over a part it was not edited from, or a state
 * undo brought back after its sheet was removed. `undefined` when the ranges hold the caches, or
 * when the part has no embedded workbook or cannot be written by the shared writer; the
 * cell-for-cell write-back then reports what it cannot write.
 */
export async function rebaseConflictingChartRanges(
	zip: JSZip,
	chart: PptxChartData,
	xml: string,
	from: string,
	inStep: string | undefined,
	resolve: ResolvePath,
	sheets: ChartDataSheets,
	readSource: ReadChartSource,
): Promise<string | undefined> {
	if (!chart.externalData || chart.externalData.isExternal) return undefined;
	try {
		const { chartSpace, issues } = readSource(xml);
		if (chartIssuesLostOnWrite(issues).length > 0) return undefined;
		const before = [from, inStep].flatMap((part) =>
			part === undefined ? [] : [readSource(part).chartSpace],
		);
		if (
			chartSpaceRangeConflicts(chartSpace, ...before).length === 0 &&
			!(await namesRemovedSheet(zip, chart, xml, inStep, resolve))
		)
			return undefined;
		return (
			await prepareWorkbook(
				zip,
				chart,
				{ xml, workbookModel: chartSpace },
				resolve,
				sheets,
				readSource,
			)
		).xml;
	} catch {
		return undefined;
	}
}
