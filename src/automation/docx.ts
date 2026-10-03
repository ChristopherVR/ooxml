import { createDocument, type Paragraph } from '../docx/model.js';
import { loadDocx } from '../docx/parse.js';
import { saveDocx } from '../docx/save.js';

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
