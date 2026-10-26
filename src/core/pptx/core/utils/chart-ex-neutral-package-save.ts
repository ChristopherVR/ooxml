// Save a supported ChartEx edit and its embedded workbook as one package update.
import type JSZip from 'jszip';
import {
	appendChartExWorkbookData,
	chartExDataChanged,
	chartExSpaceWorkbookWrites,
	parseChartExSpace,
	writeChartExSpaceOverSource,
	writeChartWorkbookUpdates,
	type ChartWorkbookWrite,
} from '../../../chart';
import {
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
} from '../../../opc/relationships';
import type { PptxChartData } from '../types';
import { planChartExNeutralSave } from './chart-ex-neutral-save';
/** A refused ChartEx save must reach the host instead of silently exporting the old chart. */
export class ChartExSaveRefusedError extends Error {}
/** Undefined means the view edit is not claimed. Claimed data failures throw before either part changes. */
export async function saveChartExPackageThroughNeutral(
	zip: JSZip,
	part: string,
	xml: string,
	before: PptxChartData,
	after: PptxChartData,
): Promise<string | undefined> {
	try {
		return await saveChartExPackage(zip, part, xml, before, after);
	} catch (error) {
		throw new ChartExSaveRefusedError(error instanceof Error ? error.message : String(error));
	}
}
async function saveChartExPackage(
	zip: JSZip,
	part: string,
	xml: string,
	before: PptxChartData,
	after: PptxChartData,
): Promise<string | undefined> {
	const edits = planChartExNeutralSave(before, after);
	if (!edits) return undefined;
	if (!edits.length) return xml;
	const parsed = parseChartExSpace(xml);
	if (parsed.issues.length) throw new Error('The ChartEx source contains unsupported content.');
	const original = parsed.chartExSpace;
	let space = original;
	for (const edit of edits) space = edit(space);
	let workbook: { part: string; bytes: Uint8Array } | undefined;
	if (chartExDataChanged(original, space)) {
		const relId = original.chartData.externalData?.relId;
		if (relId) {
			const rels = parseRelationships(await zip.file(relationshipsPartFor(part))?.async('string'));
			const rel = rels.get(relId);
			if (!rel || rel.mode === 'External' || !rel.type.endsWith('/package'))
				throw new Error('ChartEx data editing requires a local embedded workbook.');
			const target = resolvePartPath(part, rel.target);
			const bytes = await zip.file(target)?.async('uint8array');
			if (!bytes) throw new Error('The ChartEx embedded workbook is missing.');

			let prior: ChartWorkbookWrite[] | undefined, next: ChartWorkbookWrite[] | undefined;
			try {
				prior = chartExSpaceWorkbookWrites(original);
				next = chartExSpaceWorkbookWrites(space);
			} catch {
				/* Structural dimensions need a fresh sheet. */
			}
			const vectors = (writes: readonly ChartWorkbookWrite[]) =>
				writes.map((write) => [write.formula, write.isNumeric, write.values.length]);
			if (
				prior &&
				next &&
				original.chartData.data.length === space.chartData.data.length &&
				original.plotArea.series.length === space.plotArea.series.length &&
				JSON.stringify(vectors(prior)) === JSON.stringify(vectors(next))
			) {
				await writeChartWorkbookUpdates(bytes, prior, { atomic: true, requireCurrentValues: true });
				const updated = await writeChartWorkbookUpdates(bytes, next, { atomic: true });
				workbook = { part: target, bytes: updated.bytes ?? bytes };
			} else {
				const rebased = await appendChartExWorkbookData(bytes, original, space);
				space = rebased.space;
				workbook = { part: target, bytes: rebased.bytes };
			}
		} else {
			// A literal chart has no workbook to synchronize. References without a relationship are unsafe.
			if (chartExSpaceWorkbookWrites(original).length || chartExSpaceWorkbookWrites(space).length)
				throw new Error('ChartEx cell references have no embedded workbook.');
		}
	}
	const result = writeChartExSpaceOverSource(xml, space);
	if (workbook) zip.file(workbook.part, workbook.bytes);
	return result;
}
