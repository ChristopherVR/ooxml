// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Charts in a Word document: the `w:drawing` whose `a:graphicData` holds `c:chart r:id`. This
// module models the host side (placement, relationship, part name) and reads the chart part
// through the format-neutral `chart` area. The chart part, its relationships and its embedded
// workbook are never rewritten: they stay in the package untouched. The drawing is painted from
// the cached values by the shared chart painter (`chart-paint.ts`).
import {
	chartColorStyleId,
	parseChartSpace,
	readChartFormatting,
	type ChartParseIssue,
	type ChartSpace,
	type ChartStyleDefinition,
} from '../chart/index';
import { parseXml } from '../xml/index';
import { parseRelationships, relationshipsPartFor, resolvePartPath } from '../opc/index';
import type { Block } from './model';
import type { Relationship } from './package-parts';
import { forEachParagraph } from './parse-warnings';
import { getR, isElement, type XmlElement } from './xml';

/** The `a:graphicData/@uri` of a chart. */
export const CHART_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const PACKAGE_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';

/** A chart (`c:chart` in a `w:drawing`). The chart part is read-only and preserved on save. */
export interface DocxChart {
	/** `inline` (`wp:inline`) or `anchor` (`wp:anchor`, floating). */
	placement: 'inline' | 'anchor';
	/** `wp:docPr/@name` and `@id`. */
	name?: string;
	id?: string;
	/** `c:chart/@r:id`, in the part that holds the drawing. */
	relId?: string;
	/** The chart part, e.g. `word/charts/chart1.xml`, once the relationship resolves. */
	partName?: string;
	/** The parsed chart part (`c:chartSpace`); absent when the part could not be read. */
	chartSpace?: ChartSpace;
	/** Direct formatting read from the chart part (`readChartFormatting`), for painting. */
	formatting?: ChartStyleDefinition;
	/** The Office colour style id of the part's `chartColorStyle` relationship, when present. */
	colorPalette?: number;
	/** The chart title text, when the chart shows one. */
	title?: string;
	/** The embedded workbook part holding the chart data (`c:externalData`), when it is in the package. */
	workbookPartName?: string;
	/** What this model does and does not provide. */
	notice: string;
	/** Problems met resolving and reading the chart part, and what the chart model does not cover. */
	issues: ChartParseIssue[];
}

/** The honest status shown with every chart. */
export const CHART_NOTICE =
	'The chart is drawn from the values cached in its part (bar, column, line, area, pie, doughnut, scatter and radar charts; other kinds show a labelled frame). It is not editable, and the chart part and its workbook are preserved unchanged on save.';

/** Reads the host side of a chart graphic. Synchronous: resolves the relationship, reads no part. */
export function parseChartGraphic(
	graphicData: XmlElement,
	docPr: XmlElement | undefined,
	rels: ReadonlyMap<string, Relationship>,
	isAnchor: boolean,
): DocxChart {
	const element = Array.from(graphicData.childNodes).find(
		(node): node is XmlElement => isElement(node) && node.localName === 'chart',
	);
	const relId = getR(element, 'id') || undefined;
	const issues: ChartParseIssue[] = [];
	let partName: string | undefined;
	const rel = relId ? rels.get(relId) : undefined;
	if (!relId)
		issues.push({
			code: 'CHART_RELATIONSHIP_MISSING',
			message: 'The graphic has no c:chart/@r:id.',
		});
	else if (!rel || rel.mode === 'External')
		issues.push({
			code: 'CHART_RELATIONSHIP_UNRESOLVED',
			message: `Chart relationship ${relId} does not point at a part in this package.`,
		});
	else {
		if (rel.type !== CHART_RELATIONSHIP_TYPE)
			issues.push({
				code: 'CHART_RELATIONSHIP_TYPE',
				message: `Chart relationship ${relId} has type ${rel.type}, expected ${CHART_RELATIONSHIP_TYPE}.`,
			});
		partName = resolvePartPath('word/document.xml', rel.target);
	}
	const name = docPr?.getAttribute('name') || undefined;
	const id = docPr?.getAttribute('id') || undefined;
	return {
		placement: isAnchor ? 'anchor' : 'inline',
		...(name ? { name } : {}),
		...(id ? { id } : {}),
		...(relId ? { relId } : {}),
		...(partName ? { partName } : {}),
		notice: CHART_NOTICE,
		issues,
	};
}

/** Reads the chart part (and its relationships) into the model. Never throws; problems go to `issues`. */
export async function resolveChartPart(
	chart: DocxChart,
	readText: (partName: string) => Promise<string | undefined>,
): Promise<void> {
	if (!chart.partName) return;
	try {
		const xml = await readText(chart.partName);
		if (xml === undefined) {
			chart.issues.push({
				code: 'CHART_PART_MISSING',
				message: `The chart part ${chart.partName} is not in the package.`,
			});
			return;
		}
		const root = parseXml(xml, { label: 'Chart part' }).documentElement;
		const { chartSpace, issues } = parseChartSpace(root);
		chart.chartSpace = chartSpace;
		chart.issues.push(...issues);
		const title = chartSpace.autoTitleDeleted ? undefined : chartSpace.title?.text;
		if (title) chart.title = title;
		if (chartSpace.plotArea.groups.length) {
			const formatting = readChartFormatting(root);
			if (formatting) chart.formatting = formatting;
		}
		const rels = parseRelationships(await readText(relationshipsPartFor(chart.partName)));
		const dataRelId = chartSpace.externalDataRelId;
		const dataRel = dataRelId ? rels.get(dataRelId) : undefined;
		if (dataRel && dataRel.mode !== 'External' && dataRel.type === PACKAGE_RELATIONSHIP_TYPE)
			chart.workbookPartName = resolvePartPath(chart.partName, dataRel.target);
		const colorsRel = [...rels.values()].find(
			(rel) => rel.mode !== 'External' && rel.type.endsWith('/chartColorStyle'),
		);
		const colorsXml =
			colorsRel && (await readText(resolvePartPath(chart.partName, colorsRel.target)));
		const palette = colorsXml ? chartColorStyleId(colorsXml) : undefined;
		if (palette !== undefined) chart.colorPalette = palette;
	} catch (error) {
		chart.issues.push({
			code: 'CHART_PART_UNREADABLE',
			message: `The chart part ${chart.partName} could not be read: ${error instanceof Error ? error.message : String(error)}`,
		});
	}
}

/** Every chart in the given blocks, in document order. */
export function chartsIn(blocks: Block[]): DocxChart[] {
	const charts: DocxChart[] = [];
	forEachParagraph(blocks, (paragraph) => {
		for (const run of paragraph.runs) if (run.image?.chart) charts.push(run.image.chart);
	});
	return charts;
}

/** Reads the part of every chart in `blocks` (body, headers, footers, notes) through `readText`. */
export async function resolveDocumentCharts(
	blocks: Block[],
	readText: (partName: string) => Promise<string | undefined>,
): Promise<void> {
	for (const chart of chartsIn(blocks)) await resolveChartPart(chart, readText);
}
