// The package parts of an extended chart a product creates with its data in an embedded workbook:
// the `cx:chartSpace` part (`chartExN.xml`, its `cx:externalData` naming the workbook), its
// relationships, its chart style and colour style parts (which Office requires, see
// `style-part.ts`) and the new `.xlsx` under the host's embeddings folder. The data is laid out by
// `chartExDataLayout` (one column per level, then the series names) and every dimension's formula
// points at it, so Edit Data in Office opens the chart's own numbers. Shared by Word and
// PowerPoint; the host keeps its drawing frame and its relationship to the chart part.
import type JSZip from 'jszip';
import { ensureContentTypeOverride } from '../../opc/package';
import { buildRelationshipsXml, relationshipsPartFor } from '../../opc/relationships';
import { createChartWorkbookFromRows, type ChartWorkbookOptions } from '../workbook/create';
import {
	CHART_WORKBOOK_CONTENT_TYPE,
	CHART_WORKBOOK_RELATIONSHIP_TYPE,
	freePackagePartName,
	packageRelativeFolder,
	type NewChartPackageParts,
} from '../workbook/new-part';
import { chartExDataLayout } from './data-layout';
import type { ChartExSpace } from './model';
import { writeChartExSpace } from './write';
import { CHART_COLOR_STYLE_CONTENT_TYPE, CHART_COLOR_STYLE_REL } from '../color-style';
import {
	CHART_STYLE_CONTENT_TYPE,
	CHART_STYLE_REL,
	chartExColorStyleXml,
	chartExStyleXml,
} from './style-part';

export const CHART_EX_PART_CONTENT_TYPE = 'application/vnd.ms-office.chartex+xml';
/** The relationship type from a document part to an extended chart part. */
export const CHART_EX_RELATIONSHIP_TYPE =
	'http://schemas.microsoft.com/office/2014/relationships/chartEx';

export interface NewChartExPartOptions {
	/** The chart part's folder, with its trailing slash (`word/charts/`, `ppt/charts/`). */
	chartFolder: string;
	/** The embedded workbook's folder, with its trailing slash (`word/embeddings/`). */
	embeddingsFolder: string;
	/** The worksheet name; `Sheet1` by default. */
	sheetName?: string;
	/** How the new workbook is written (the host theme's body font, its ZIP entry date). */
	workbook?: ChartWorkbookOptions;
}

/**
 * Stages the parts of a new extended chart over `space`'s cached data (literal data, as
 * `createChartExSpace` makes it); the caller adds the content types. Returns the part names and
 * the model as written (formulas pointing at the new workbook).
 */
export function stageChartExPartWithWorkbook(
	zip: JSZip,
	space: ChartExSpace,
	options: NewChartExPartOptions,
): NewChartPackageParts & { space: ChartExSpace } {
	const { chartFolder, embeddingsFolder } = options;
	const sheetName = options.sheetName ?? 'Sheet1';
	const partName = freePackagePartName(zip, chartFolder, 'chartEx', 'xml', 1);
	const workbook = freePackagePartName(
		zip,
		embeddingsFolder,
		'Microsoft_Excel_Worksheet',
		'xlsx',
		0,
	);
	const relId = 'rId1';
	const layout = chartExDataLayout(space, sheetName);
	const written: ChartExSpace = {
		...layout.space,
		chartData: { ...layout.space.chartData, externalData: { relId, autoUpdate: false } },
	};
	zip.file(partName, writeChartExSpace(written));
	const style = freePackagePartName(zip, chartFolder, 'style', 'xml', 1);
	zip.file(style, chartExStyleXml());
	const colors = freePackagePartName(zip, chartFolder, 'colors', 'xml', 1);
	zip.file(colors, chartExColorStyleXml());
	const target = `${packageRelativeFolder(chartFolder, embeddingsFolder)}${workbook.slice(embeddingsFolder.length)}`;
	const local = (part: string) => part.slice(chartFolder.length);
	zip.file(
		relationshipsPartFor(partName),
		buildRelationshipsXml(
			new Map([
				[relId, { type: CHART_WORKBOOK_RELATIONSHIP_TYPE, target, mode: 'Internal' as const }],
				['rId2', { type: CHART_STYLE_REL, target: local(style), mode: 'Internal' as const }],
				['rId3', { type: CHART_COLOR_STYLE_REL, target: local(colors), mode: 'Internal' as const }],
			]),
		),
	);
	zip.file(workbook, createChartWorkbookFromRows(layout.rows, sheetName, options.workbook));
	return {
		partName,
		workbookPartName: workbook,
		workbookRelationshipId: relId,
		contentTypes: [
			[partName, CHART_EX_PART_CONTENT_TYPE],
			[style, CHART_STYLE_CONTENT_TYPE],
			[colors, CHART_COLOR_STYLE_CONTENT_TYPE],
			[workbook, CHART_WORKBOOK_CONTENT_TYPE],
		],
		space: written,
	};
}

/** {@link stageChartExPartWithWorkbook} with the content-type overrides added. */
export async function writeChartExPartWithWorkbook(
	zip: JSZip,
	space: ChartExSpace,
	options: NewChartExPartOptions,
): Promise<NewChartPackageParts & { space: ChartExSpace }> {
	const parts = stageChartExPartWithWorkbook(zip, space, options);
	for (const [partName, contentType] of parts.contentTypes)
		await ensureContentTypeOverride(zip, partName, contentType);
	return parts;
}
