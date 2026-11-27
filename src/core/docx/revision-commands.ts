// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph, Revision } from './model';
import { restoreRunFormatting } from './restore-run-format';
import { restoreParagraphFormatting } from './restore-paragraph-format';
import { documentBlockLists } from './document-paragraphs';
import { isEmptyParagraph } from './paragraph-content';

export interface RevisionEntry {
	id: string;
	kind: Revision['kind'];
	author: string;
	date?: string;
	paragraphId: string;
	/** Present for run-level revisions; absent for paragraph mark/format revisions. */
	runIndex?: number;
	/** The move a `moveFrom`/`moveTo` revision belongs to; both sides resolve together. */
	move?: { name: string; rangeId?: string };
	/** A whole-row insertion or deletion (`w:trPr/w:ins|w:del`): the table and its row. */
	row?: { tableId: string; index: number };
}

function paragraphsOf(blocks: Block[]): Paragraph[] {
	const paragraphs: Paragraph[] = [];
	for (const block of blocks)
		if (block.type === 'paragraph') paragraphs.push(block);
		else for (const row of block.rows) for (const cell of row) paragraphs.push(...cell.paragraphs);
	return paragraphs;
}

function documentParagraphs(model: DocumentModel): Paragraph[] {
	return documentBlockLists(model).flatMap(paragraphsOf);
}

/** Row revisions of a story's tables, keyed by the first paragraph of the row. */
function rowEntries(blocks: Block[]): Map<string, RevisionEntry> {
	const entries = new Map<string, RevisionEntry>();
	for (const block of blocks)
		if (block.type === 'table')
			block.rows.forEach((row, index) => {
				const revision = block.rowProperties?.[index]?.revision;
				const paragraphId = row[0]?.paragraphs[0]?.id;
				if (revision && paragraphId)
					entries.set(paragraphId, {
						...revision,
						paragraphId,
						row: { tableId: block.id, index },
					});
			});
	return entries;
}

/** Lists every row/run/paragraph revision in document order, for navigation and enumeration. */
export function listRevisions(model: DocumentModel): RevisionEntry[] {
	const entries: RevisionEntry[] = [];
	const rows = new Map(documentBlockLists(model).flatMap((blocks) => [...rowEntries(blocks)]));
	for (const paragraph of documentParagraphs(model)) {
		const row = rows.get(paragraph.id);
		if (row) entries.push(row);
		if (paragraph.markRevision)
			entries.push({ ...paragraph.markRevision, paragraphId: paragraph.id });
		if (paragraph.formatRevision)
			entries.push({ ...paragraph.formatRevision, paragraphId: paragraph.id });
		paragraph.runs.forEach((run, runIndex) => {
			if (run.revision) entries.push({ ...run.revision, paragraphId: paragraph.id, runIndex });
			if (run.formatRevision)
				entries.push({ ...run.formatRevision, paragraphId: paragraph.id, runIndex });
		});
	}
	return entries;
}

export function findRevision(model: DocumentModel, id: string): RevisionEntry | undefined {
	return listRevisions(model).find((entry) => entry.id === id);
}

function mergeIntoNext(blocks: Block[], paragraphId: string): Block[] {
	const index = blocks.findIndex((b) => b.type === 'paragraph' && b.id === paragraphId);
	const next = blocks[index + 1];
	// An emptied paragraph before a table merges with nothing: it goes, as Word removes it.
	if (index >= 0 && next?.type === 'table' && isEmptyParagraph(blocks[index] as Paragraph))
		return [...blocks.slice(0, index), ...blocks.slice(index + 1)];
	if (index < 0 || !next || next.type !== 'paragraph')
		throw new Error(
			`Cannot resolve the paragraph mark revision on paragraph ${paragraphId}: there is no following paragraph to merge with (merging across a table boundary is not supported).`,
		);
	const current = blocks[index] as Paragraph;
	const merged: Paragraph = { ...next, runs: [...current.runs, ...next.runs] };
	return [...blocks.slice(0, index), merged, ...blocks.slice(index + 2)];
}

/** All runs carrying one revision share the same wrapper `id`; ordered indices for splicing. */
function matchingRunIndices(paragraph: Paragraph, id: string): number[] {
	return paragraph.runs.reduce<number[]>((indices, run, index) => {
		if (run.revision?.id === id || run.formatRevision?.id === id) indices.push(index);
		return indices;
	}, []);
}
function updateRuns(
	model: DocumentModel,
	id: string,
	apply: (paragraph: Paragraph, runIndex: number) => void,
): DocumentModel {
	const next = structuredClone(model);
	for (const paragraph of documentParagraphs(next)) {
		const indices = matchingRunIndices(paragraph, id);
		if (!indices.length) continue;
		for (const index of [...indices].sort((a, b) => b - a)) apply(paragraph, index);
		if (!paragraph.runs.length) paragraph.runs.push({ text: '' });
		return next;
	}
	throw new Error(`No run revision with id ${id} was found.`);
}

/** Resolve a paragraph boundary using the following paragraph's formatting when merging. */
export function resolveParagraphMark(
	model: DocumentModel,
	id: string,
	keepBreak: boolean,
): DocumentModel {
	const next = structuredClone(model);
	const paragraph = documentParagraphs(next).find((p) => p.markRevision?.id === id);
	if (!paragraph) throw new Error(`No paragraph mark revision with id ${id} was found.`);
	if (keepBreak) {
		delete paragraph.markRevision;
		return next;
	}
	for (const blocks of documentBlockLists(next)) {
		const lists: Block[][] = [
			blocks,
			...blocks.flatMap((block) =>
				block.type === 'table'
					? block.rows.flatMap((row) => row.map((cell) => cell.paragraphs))
					: [],
			),
		];
		for (const [index, list] of lists.entries()) {
			if (!list.includes(paragraph)) continue;
			// A cell's last mark ends the cell: Word never merges it away, so the break stays.
			if (index > 0 && list.at(-1) === paragraph) {
				delete paragraph.markRevision;
				return next;
			}
			const merged = mergeIntoNext(list, paragraph.id);
			list.splice(0, list.length, ...merged);
			return next;
		}
	}
	throw new Error(`Cannot locate the story containing paragraph ${paragraph.id}.`);
}

/** Revision ids resolved together with `id`: every side of the same move, or just `id`. */
export function linkedRevisionIds(model: DocumentModel, id: string): string[] {
	const entries = listRevisions(model);
	const move = entries.find((entry) => entry.id === id)?.move;
	if (!move) return [id];
	const ids = entries
		.filter(
			(entry) =>
				(entry.kind === 'moveFrom' || entry.kind === 'moveTo') && entry.move?.name === move.name,
		)
		.map((entry) => entry.id);
	return [...new Set([id, ...ids])];
}

function resolveLinked(
	model: DocumentModel,
	id: string,
	resolve: (model: DocumentModel, id: string) => DocumentModel,
): DocumentModel {
	if (!findRevision(model, id)) throw new Error(`No revision with id ${id} was found.`);
	let current = model;
	for (const linked of linkedRevisionIds(model, id))
		while (findRevision(current, linked)) current = resolve(current, linked);
	return current;
}

/** Accepts one revision (both sides of a move): keeps insertions, drops deletions, clears format markers. */
export const acceptRevision = (model: DocumentModel, id: string): DocumentModel =>
	resolveLinked(model, id, acceptOne);
/** Rejects one revision (both sides of a move): drops insertions, restores deletions. */
export const rejectRevision = (model: DocumentModel, id: string): DocumentModel =>
	resolveLinked(model, id, rejectOne);

function clearRunRevision(paragraph: Paragraph, runIndex: number): void {
	const run = paragraph.runs[runIndex];
	if (run) delete run.revision;
}

/** Keeps (`remove` false) or removes a tracked row; a table left without rows goes too. */
function resolveRow(model: DocumentModel, id: string, remove: boolean): DocumentModel {
	const next = structuredClone(model);
	for (const blocks of documentBlockLists(next))
		for (const [blockIndex, table] of blocks.entries()) {
			if (table.type !== 'table') continue;
			const index = table.rowProperties?.findIndex((row) => row.revision?.id === id) ?? -1;
			if (index < 0) continue;
			if (!remove) {
				delete table.rowProperties![index]!.revision;
				return next;
			}
			table.rows.splice(index, 1);
			table.rowProperties!.splice(index, 1);
			if (!table.rows.length) blocks.splice(blockIndex, 1);
			return next;
		}
	throw new Error(`No table row revision with id ${id} was found.`);
}

function acceptOne(model: DocumentModel, id: string): DocumentModel {
	const entry = findRevision(model, id);
	if (!entry) throw new Error(`No revision with id ${id} was found.`);
	if (entry.row) return resolveRow(model, id, entry.kind === 'delete');
	if (entry.kind === 'formatChange')
		return updateRuns(model, id, (paragraph, runIndex) => {
			const run = paragraph.runs[runIndex]!;
			if (run.formatRevision?.id === id) delete run.formatRevision;
			else delete run.revision;
		});
	if (entry.kind === 'paragraphChange') {
		const next = structuredClone(model);
		const paragraph = documentParagraphs(next).find((p) => p.formatRevision?.id === id)!;
		delete paragraph.formatRevision;
		return next;
	}
	if (entry.runIndex === undefined)
		return resolveParagraphMark(model, id, entry.kind === 'insert' || entry.kind === 'moveTo');
	if (entry.kind === 'insert' || entry.kind === 'moveTo')
		return updateRuns(model, id, (paragraph, runIndex) => {
			clearRunRevision(paragraph, runIndex);
		});
	return updateRuns(model, id, (paragraph, runIndex) => {
		paragraph.runs.splice(runIndex, 1);
	});
}

function rejectOne(model: DocumentModel, id: string): DocumentModel {
	const entry = findRevision(model, id);
	if (!entry) throw new Error(`No revision with id ${id} was found.`);
	if (entry.row) return resolveRow(model, id, entry.kind === 'insert');
	if (entry.kind === 'formatChange')
		return updateRuns(model, id, (paragraph, index) =>
			restoreRunFormatting(paragraph.runs[index]!),
		);
	if (entry.kind === 'paragraphChange') {
		const next = structuredClone(model);
		const paragraph = documentParagraphs(next).find((item) => item.formatRevision?.id === id)!;
		restoreParagraphFormatting(paragraph);
		return next;
	}
	if (entry.runIndex === undefined)
		return resolveParagraphMark(model, id, entry.kind === 'delete' || entry.kind === 'moveFrom');
	if (entry.kind === 'insert' || entry.kind === 'moveTo')
		return updateRuns(model, id, (paragraph, runIndex) => {
			paragraph.runs.splice(runIndex, 1);
		});
	return updateRuns(model, id, (paragraph, runIndex) => {
		clearRunRevision(paragraph, runIndex);
	});
}

function applyAll(
	model: DocumentModel,
	apply: (model: DocumentModel, id: string) => DocumentModel,
): DocumentModel {
	let current = model;
	// Whole rows first: a removed row takes the revisions inside it along.
	const entries = listRevisions(model).reverse();
	for (const entry of [...entries.filter((item) => item.row), ...entries]) {
		if (!findRevision(current, entry.id)) continue;
		current = apply(current, entry.id);
	}
	return current;
}

export const acceptAllRevisions = (model: DocumentModel): DocumentModel =>
	applyAll(model, acceptRevision);
export const rejectAllRevisions = (model: DocumentModel): DocumentModel =>
	applyAll(model, rejectRevision);
