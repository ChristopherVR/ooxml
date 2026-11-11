/**
 * The package parts of a chart a product creates with its data in an embedded workbook, the way
 * Word and PowerPoint write Insert > Chart: the chart part (with `c:externalData` naming the
 * workbook), its relationships (the workbook's `package` relationship plus any the host adds, such
 * as a colour-style part) and the new `.xlsx` under the host's embeddings folder. Shared by Word
 * (`docx/chart-save.ts`) and PowerPoint (`pptx/core/core/runtime`); the host keeps its own drawing
 * frame and the relationship from its document part to the chart.
 *
 * {@link stageChartPartWithWorkbook} stages every part at once (the workbook as a pending promise,
 * which JSZip resolves when the package is generated), so a host that names further parts before
 * awaiting sees them taken; the host registers the returned content types. {@link writeChartPartWithWorkbook}
 * also adds the content-type overrides.
 *
 * @module chart/workbook/new-part
 */
import type JSZip from 'jszip';
import { ensureContentTypeOverride } from '../../opc/package';
import {
	buildRelationshipsXml,
	nextRelationshipId,
	relationshipsPartFor,
	type RelationshipInput,
} from '../../opc/relationships';
import { chartSpaceDataSheet } from '../edit/data-sheet';
import type { ChartSpace } from '../model';
import { writeChartSpace } from '../write-space';
import { parseChartFormulaRange } from './cellref';
import { createChartWorkbook, createChartWorkbookFromRows } from './create';

export const CHART_PART_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const CHART_WORKBOOK_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const CHART_WORKBOOK_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';

/** The first `<folder><stem><n>.<ext>` the package does not hold (`n` from `start`, 0 written as none). */
export function freePackagePartName(
	zip: JSZip,
	folder: string,
	stem: string,
	ext: string,
	start: number,
): string {
	const taken = new Set(Object.keys(zip.files).map((name) => name.toLowerCase()));
	for (let n = start; ; n++) {
		const name = `${folder}${stem}${n || ''}.${ext}`;
		if (!taken.has(name.toLowerCase())) return name;
	}
}

export interface NewChartPartOptions {
	/** The chart part's folder, with its trailing slash (`word/charts/`, `ppt/charts/`). */
	chartFolder: string;
	/** The embedded workbook's folder, with its trailing slash (`word/embeddings/`). */
	embeddingsFolder: string;
	/** The chart part's name when the host picked it; else the first free `chartN.xml` (from 1). */
	partName?: string;
	/**
	 * Further relationships of the chart part, keyed by id (a colour-style part). The workbook
	 * takes the first `rIdN` none of them uses, and is written first.
	 */
	relationships?: ReadonlyMap<string, RelationshipInput>;
	/**
	 * The worksheet rows (from row 1, `null` an empty cell) when the chart's formulas are not the
	 * standard layout of its data sheet (`chartSpaceDataSheet`), which is written otherwise.
	 */
	rows?: readonly (readonly (string | number | null)[])[];
	/** The sheet name; by default the one the first series' values formula names, else `Sheet1`. */
	sheetName?: string;
}

/** The parts written for a new chart. */
export interface NewChartPackageParts {
	partName: string;
	workbookPartName: string;
	/** The id of the chart part's relationship to the workbook (`c:externalData/@r:id`). */
	workbookRelationshipId: string;
	/** The content-type overrides the parts need, as `[partName, contentType]`. */
	contentTypes: [string, string][];
}

/** Stages the chart part, its relationships and its workbook; the caller adds the content types. */
export function stageChartPartWithWorkbook(
	zip: JSZip,
	space: ChartSpace,
	options: NewChartPartOptions,
): NewChartPackageParts {
	const { chartFolder, embeddingsFolder } = options;
	const partName = options.partName ?? freePackagePartName(zip, chartFolder, 'chart', 'xml', 1);
	const workbook = freePackagePartName(
		zip,
		embeddingsFolder,
		'Microsoft_Excel_Worksheet',
		'xlsx',
		0,
	);
	const firstFormula = space.plotArea.groups[0]?.series[0]?.values?.formula;
	const sheetName =
		options.sheetName ??
		((firstFormula && parseChartFormulaRange(firstFormula)?.sheetName) || undefined);
	const extra = options.relationships ?? new Map<string, RelationshipInput>();
	const relId = nextRelationshipId(new Set(extra.keys()));
	zip.file(
		partName,
		writeChartSpace({ ...space, externalDataRelId: relId, externalDataAutoUpdate: false }),
	);
	const target = `${relativeFolder(chartFolder, embeddingsFolder)}${workbook.slice(embeddingsFolder.length)}`;
	zip.file(
		relationshipsPartFor(partName),
		buildRelationshipsXml(
			new Map<string, RelationshipInput>([
				[relId, { type: CHART_WORKBOOK_RELATIONSHIP_TYPE, target, mode: 'Internal' }],
				...extra,
			]),
		),
	);
	zip.file(
		workbook,
		options.rows
			? createChartWorkbookFromRows(options.rows, sheetName)
			: createChartWorkbook(chartSpaceDataSheet(space), sheetName),
	);
	return {
		partName,
		workbookPartName: workbook,
		workbookRelationshipId: relId,
		contentTypes: [
			[partName, CHART_PART_CONTENT_TYPE],
			[workbook, CHART_WORKBOOK_CONTENT_TYPE],
		],
	};
}

/** {@link stageChartPartWithWorkbook} with the content-type overrides added. */
export async function writeChartPartWithWorkbook(
	zip: JSZip,
	space: ChartSpace,
	options: NewChartPartOptions,
): Promise<NewChartPackageParts> {
	const parts = stageChartPartWithWorkbook(zip, space, options);
	for (const [partName, contentType] of parts.contentTypes)
		await ensureContentTypeOverride(zip, partName, contentType);
	return parts;
}

/** The relative path from `from/` to `to/` (`word/charts/` to `word/embeddings/` is `../embeddings/`). */
function relativeFolder(from: string, to: string): string {
	const a = from.split('/').filter(Boolean);
	const b = to.split('/').filter(Boolean);
	let common = 0;
	while (common < a.length && common < b.length && a[common] === b[common]) common++;
	return `${'../'.repeat(a.length - common)}${b.slice(common).join('/')}${b.length > common ? '/' : ''}`;
}
