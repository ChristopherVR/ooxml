// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Saving Word charts. An edited chart (its `chartSpace` differs from the loaded one) is written
// back to its own part with the shared chart writer, and the embedded workbook its
// `c:externalData` names is brought in line with the new caches cell for cell
// (`chart/workbook`), so Word's Edit Data opens the edited numbers. A chart inserted in this
// session gets a new chart part, its relationships, a new embedded workbook in the standard
// layout and content types, the parts Word itself writes for Insert > Chart (minus the optional
// chart style and colour parts, which Word does not require). Charts that were not edited are
// never touched, so their parts stay byte for byte.
import type JSZip from 'jszip';
import { applyChartExPartEdits } from './chart-ex-save';
import {
	writeChartSpaceOverSource,
	chartSpaceWorkbookWrites,
	saveChartWorkbookUpdates,
	writeChartPartWithWorkbook,
	type ChartSpace,
	type ChartWorkbookIssue,
} from '../chart/index';
import { chartsIn, documentDrawingBlocks, type DocxChart } from './chart';
import type { DocumentModel } from './model';
import { isNewDocxChart } from './chart-edit';

/**
 * Writes the parts of a chart inserted in this session: the chart part, its relationships, its
 * embedded workbook and their content types. Returns the new chart part's name.
 */
export async function writeNewChartPart(zip: JSZip, space: ChartSpace): Promise<string> {
	const { partName } = await writeChartPartWithWorkbook(zip, space, {
		chartFolder: 'word/charts/',
		embeddingsFolder: 'word/embeddings/',
	});
	return partName;
}

/** Writes one edited chart part and updates its workbook; problems go to `report`. */
async function saveEditedChart(
	zip: JSZip,
	chart: DocxChart & { partName: string; chartSpace: ChartSpace },
	report: (issue: ChartWorkbookIssue) => void,
): Promise<void> {
	const source = await zip.file(chart.partName)?.async('string');
	if (!source) throw new Error('The source chart part is missing.');
	zip.file(chart.partName, writeChartSpaceOverSource(source, chart.chartSpace));
	await saveChartWorkbookUpdates(
		{
			read: async (partName) => zip.file(partName)?.async('uint8array'),
			write: (partName, bytes) => {
				zip.file(partName, bytes);
			},
		},
		chart.partName,
		chart.chartSpace.externalDataRelId
			? chart.workbookPartName
				? { partName: chart.workbookPartName }
				: { external: true }
			: undefined,
		chartSpaceWorkbookWrites(chart.chartSpace),
		{ report },
	);
}

/**
 * Saves every chart of `model` whose part was loaded and whose model was edited since `base`.
 * Returns the workbook issues met (data saved to the chart cache only).
 */
export async function applyChartPartEdits(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel | undefined,
): Promise<ChartWorkbookIssue[]> {
	const issues: ChartWorkbookIssue[] = [];
	await applyChartExPartEdits(zip, model, base);
	if (!base) return issues;
	const loaded = new Map<string, string>();
	for (const chart of chartsIn(documentDrawingBlocks(base)))
		if (chart.partName && chart.chartSpace)
			loaded.set(chart.partName, JSON.stringify(chart.chartSpace));
	const saved = new Set<string>();
	for (const chart of chartsIn(documentDrawingBlocks(model))) {
		const { partName, chartSpace } = chart;
		if (!partName || !chartSpace || isNewDocxChart(chart) || saved.has(partName)) continue;
		const before = loaded.get(partName);
		if (before === undefined || before === JSON.stringify(chartSpace)) continue;
		saved.add(partName);
		await saveEditedChart(zip, { ...chart, partName, chartSpace }, (issue) => issues.push(issue));
	}
	return issues;
}
