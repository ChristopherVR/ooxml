import { createDocument, type Paragraph } from '../docx/model';
import { loadDocx } from '../docx/parse';
import { saveDocx } from '../docx/save';

/** Headless Word operations, shared by MCP and other automation adapters. */
export async function inspectDocx(bytes: Uint8Array) {
	const { model } = await loadDocx(bytes);
	return { blocks: model.blocks, properties: model.properties, warnings: model.warnings };
}

export async function createDocx(paragraphs: readonly string[]): Promise<Uint8Array> {
	const model = createDocument();
	if (paragraphs.length)
		model.blocks = paragraphs.map((text, i) => ({
			type: 'paragraph',
			id: `p${i + 1}`,
			runs: [{ text }],
		}));
	return saveDocx(model);
}

/** Edits a single ordinary text run without flattening paragraph formatting or fields. */
export async function setDocxRunText(
	bytes: Uint8Array,
	paragraphId: string,
	runIndex: number,
	text: string,
) {
	const loaded = await loadDocx(bytes);
	const paragraphs: Paragraph[] = loaded.model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
	const paragraph = paragraphs.find((item) => item.id === paragraphId);
	const run =
		Number.isSafeInteger(runIndex) && runIndex >= 0 ? paragraph?.runs[runIndex] : undefined;
	if (!run) throw new Error('Paragraph or run does not exist');
	if (
		loaded.model.trackChanges ||
		paragraph?.markRevision ||
		paragraph?.formatRevision ||
		run.equation ||
		run.image ||
		run.revision ||
		run.field ||
		run.fieldChar ||
		run.fieldCode !== undefined
	)
		throw new Error('Only ordinary text runs can be edited by this operation');
	run.text = text;
	return { bytes: await loaded.save(), warnings: loaded.model.warnings };
}

export interface DocxRunSuggestion {
	paragraphId: string;
	runIndex: number;
	text: string;
}

/**
 * Proposes run edits as Word tracked changes: each run's current text becomes a deletion and the
 * new text an insertion with the same formatting, both by `author`. Word, and the suite's Word
 * editor, show them as redlines the user accepts or rejects one by one. Run indexes refer to the
 * document as it is, before any of these suggestions.
 */
export async function suggestDocxRunText(
	bytes: Uint8Array,
	suggestions: readonly DocxRunSuggestion[],
	{
		author = 'Assistant',
		date = new Date().toISOString(),
	}: { author?: string; date?: string } = {},
) {
	const loaded = await loadDocx(bytes);
	const paragraphs: Paragraph[] = loaded.model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
	let next = 0;
	const revision = (kind: 'insert' | 'delete') => ({
		kind,
		author,
		date: date.replace(/\.\d+Z$/u, 'Z'),
		id: `assistant-${next++}`,
	});
	// Later runs first, so inserting a run never shifts an index still to be edited.
	const ordered = [...suggestions].sort((a, b) => b.runIndex - a.runIndex);
	for (const suggestion of ordered) {
		const paragraph = paragraphs.find((item) => item.id === suggestion.paragraphId);
		const index = suggestion.runIndex;
		const run = Number.isSafeInteger(index) && index >= 0 ? paragraph?.runs[index] : undefined;
		if (!paragraph || !run) throw new Error('Paragraph or run does not exist');
		if (
			run.equation ||
			run.image ||
			run.revision ||
			run.field ||
			run.fieldChar ||
			run.fieldCode !== undefined
		)
			throw new Error('Only ordinary text runs can be suggested');
		if (run.text === suggestion.text) continue;
		const inserted = {
			...structuredClone(run),
			text: suggestion.text,
			revision: revision('insert'),
		};
		paragraph.runs.splice(index, 1, { ...run, revision: revision('delete') }, inserted);
	}
	return { bytes: await loaded.save(), warnings: loaded.model.warnings };
}
