// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word requires decimal revision ids (`w:ins/@w:id`, `w:del/@w:id`, …; ST_DecimalNumber).
import type { DocumentModel, Paragraph, Revision, Table } from './model';
import {
	documentBlockLists,
	mapBlockParagraphs,
	mapDocumentParagraphs,
} from './document-paragraphs';

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
			collect(paragraph.markInsertion);
			collect(paragraph.formatRevision);
			for (const run of paragraph.runs) {
				collect(run.revision);
				collect(run.formatRevision);
			}
			return paragraph;
		});
	for (const table of documentTables(model))
		for (const row of table.rowProperties ?? []) collect(row.revision);
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
		...(block.markInsertion ? { markInsertion: rename(block.markInsertion) } : {}),
		...(block.formatRevision ? { formatRevision: rename(block.formatRevision) } : {}),
		runs: block.runs.map((run) => ({
			...run,
			...(run.revision ? { revision: rename(run.revision) } : {}),
			...(run.formatRevision ? { formatRevision: rename(run.formatRevision) } : {}),
		})),
	});
	const result = mapDocumentParagraphs(model, paragraph);
	// mapDocumentParagraphs copies every table, so its row properties can be replaced in place.
	for (const table of documentTables(result))
		if (table.rowProperties?.some((row) => row.revision))
			table.rowProperties = table.rowProperties.map((row) =>
				row.revision ? { ...row, revision: rename(row.revision) } : row,
			);
	return result;
}

/** Every table of every story, with row revisions to number. */
function documentTables(model: DocumentModel): Table[] {
	return documentBlockLists(model).flatMap((blocks) =>
		blocks.filter((block): block is Table => block.type === 'table'),
	);
}

/** The largest decimal `w:id` in a part's XML, so new ids never collide with existing markers. */
export function maxWordId(xml: string | undefined): number {
	let max = -1;
	for (const match of xml?.matchAll(/\bw:id="(\d{1,9})"/g) ?? [])
		max = Math.max(max, Number(match[1]));
	return max;
}
