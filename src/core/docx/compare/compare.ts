import type { Block, DocumentModel } from '../model';
import { mapDocumentParagraphs } from '../document-paragraphs';
import { loadDocx } from '../parse';
import {
	acceptAllRevisions,
	acceptRevision,
	findRevision,
	listRevisions,
} from '../revision-commands';
import { createContext, revisionDate } from './context';
import { compareStory } from './story';
import { tableLayout } from './table';
import {
	UNCOMPARED_DOCUMENT_PARTS,
	type CompareDocumentsOptions,
	type CompareDocumentsResult,
} from './types';

/** Keeps a paragraph break whose tracked mark cannot be resolved (the end of a story). */
function keepMark(model: DocumentModel, id: string): DocumentModel {
	return mapDocumentParagraphs(model, (paragraph) => {
		if (paragraph.markRevision?.id !== id) return paragraph;
		const { markRevision: _resolved, ...rest } = paragraph;
		return rest;
	});
}

/**
 * Every tracked change of a document accepted, as Word does before comparing ("Word will consider
 * these changes accepted for the purpose of the comparison").
 */
export function acceptForComparison(model: DocumentModel, warnings: string[]): DocumentModel {
	if (!listRevisions(model).length) return model;
	try {
		return acceptAllRevisions(model);
	} catch {
		let current = model;
		for (const entry of listRevisions(model).reverse()) {
			if (!findRevision(current, entry.id)) continue;
			try {
				current = acceptRevision(current, entry.id);
			} catch {
				current = keepMark(current, entry.id);
				const message =
					'A tracked paragraph mark that could not be accepted was kept before comparing.';
				if (!warnings.includes(message)) warnings.push(message);
			}
		}
		return current;
	}
}

/** Ids of the body paragraphs (table cells included) that hold a picture or drawing. */
function picturedParagraphs(model: DocumentModel): string[] {
	const ids: string[] = [];
	const visit = (blocks: Block[]) => {
		for (const block of blocks)
			if (block.type === 'table')
				for (const row of block.rows) for (const cell of row) visit(cell.paragraphs);
			else if (block.runs.some((run) => run.image)) ids.push(block.id);
	};
	visit(model.blocks);
	return ids;
}

/** Paragraphs compared again as whole insertions before a save failure is final. */
const MAX_WHOLE_PARAGRAPHS = 64;

const toBytes = (input: Uint8Array | ArrayBuffer): Uint8Array =>
	input instanceof Uint8Array ? input : new Uint8Array(input);

/**
 * Word's Review > Compare: a new document whose content is the revised document, with what
 * changed since the original as tracked changes by `author`. Tracked changes already in either
 * document are accepted first. Paragraphs with the same text are aligned first; the paragraphs
 * between them are compared word by word (or character by character), paragraph marks included,
 * so a paragraph that only changed shows its changed words and a whole paragraph shows inserted or
 * deleted with its mark. Tables are compared row by row and cell by cell; formatting changes of
 * unchanged text are `w:rPrChange` revisions. See {@link UNCOMPARED_DOCUMENT_PARTS} for what is not compared.
 */
export async function compareDocuments(
	original: Uint8Array | ArrayBuffer,
	revised: Uint8Array | ArrayBuffer,
	options: CompareDocumentsOptions,
): Promise<CompareDocumentsResult> {
	const [originalDocument, revisedDocument] = await Promise.all([
		loadDocx(toBytes(original)),
		loadDocx(toBytes(revised)),
	]);
	const accepted: string[] = [];
	const before = acceptForComparison(originalDocument.model, accepted);
	const after = acceptForComparison(revisedDocument.model, accepted);
	// The writer refuses to re-split a paragraph whose runs carry properties it cannot rebuild;
	// such a paragraph is compared again shown inserted whole, which keeps its runs as they are.
	const sourceLayouts = new Map(
		revisedDocument.model.blocks.flatMap((block) =>
			block.type === 'table' ? [[block.id, tableLayout(block)] as const] : [],
		),
	);
	const whole = new Set<string>();
	for (;;) {
		const warnings = [...accepted];
		if (whole.size)
			warnings.push(
				'Paragraphs whose runs carry formatting the writer cannot split are shown inserted whole, with the original text deleted.',
			);
		const context = createContext({
			author: options.author,
			date: revisionDate(options.date),
			granularity: options.granularity ?? 'word',
			compareFormatting: options.compareFormatting ?? true,
			revised: after,
			sourceLayouts,
			whole,
			warnings,
		});
		const blocks = compareStory(before.blocks, after.blocks, context);
		const result: DocumentModel = structuredClone({ ...after, blocks });
		let bytes: Uint8Array;
		try {
			bytes = await revisedDocument.save(result);
		} catch (error) {
			const id = /^Cannot edit paragraph (\S+?):/.exec((error as Error).message)?.[1];
			// Other refusals (a picture the writer cannot replace in place) name no paragraph:
			// every paragraph with a picture is then shown whole, once.
			const added = id ? [id] : picturedParagraphs(after).filter((item) => !whole.has(item));
			if (!added.length || whole.has(added[0]!) || whole.size >= MAX_WHOLE_PARAGRAPHS) throw error;
			for (const item of added) whole.add(item);
			continue;
		}
		const reloaded = await loadDocx(bytes);
		return {
			bytes,
			model: reloaded.model,
			summary: context.summary,
			warnings,
			notCompared: UNCOMPARED_DOCUMENT_PARTS,
		};
	}
}
