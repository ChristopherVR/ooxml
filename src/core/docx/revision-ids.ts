// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word requires decimal revision ids (`w:ins/@w:id`, `w:del/@w:id`, …; ST_DecimalNumber).
import type { DocumentModel, Paragraph, Revision } from './model.js';
import {
	documentBlockLists,
	mapBlockParagraphs,
	mapDocumentParagraphs,
} from './document-paragraphs.js';

const DECIMAL_ID = /^\d{1,9}$/;

/**
 * Renumbers editor-minted revision ids (such as `dve-rev-…`) to unused decimals, consistently
 * wherever a revision appears, after the largest id in the source XML (`reservedMax`) and in the
 * model. Returns the model unchanged when every id is already decimal.
 */
export function numberRevisionIds(model: DocumentModel, reservedMax = -1): DocumentModel {
	const ids = new Set<string>();
	const collect = (revision: Revision | undefined) => revision && ids.add(revision.id);
	for (const blocks of documentBlockLists(model))
		mapBlockParagraphs(blocks, (paragraph) => {
			collect(paragraph.markRevision);
			collect(paragraph.formatRevision);
			for (const run of paragraph.runs) {
				collect(run.revision);
				collect(run.formatRevision);
			}
			return paragraph;
		});
	const invalid = [...ids].filter((id) => !DECIMAL_ID.test(id));
	if (!invalid.length) return model;
	let next =
		Math.max(
			reservedMax,
			...[...ids].filter((id) => DECIMAL_ID.test(id)).map(Number),
			...(model.comments ?? []).map((comment) => Number(comment.id)).filter(Number.isSafeInteger),
		) + 1;
	const renamed = new Map(invalid.map((id) => [id, String(next++)]));
	const rename = <T extends Revision | undefined>(revision: T): T =>
		revision && renamed.has(revision.id)
			? { ...revision, id: renamed.get(revision.id)! }
			: revision;
	const paragraph = (block: Paragraph): Paragraph => ({
		...block,
		...(block.markRevision ? { markRevision: rename(block.markRevision) } : {}),
		...(block.formatRevision ? { formatRevision: rename(block.formatRevision) } : {}),
		runs: block.runs.map((run) => ({
			...run,
			...(run.revision ? { revision: rename(run.revision) } : {}),
			...(run.formatRevision ? { formatRevision: rename(run.formatRevision) } : {}),
		})),
	});
	return mapDocumentParagraphs(model, paragraph);
}

/** The largest decimal `w:id` in a part's XML, so new ids never collide with existing markers. */
export function maxWordId(xml: string | undefined): number {
	let max = -1;
	for (const match of xml?.matchAll(/\bw:id="(\d{1,9})"/g) ?? [])
		max = Math.max(max, Number(match[1]));
	return max;
}
