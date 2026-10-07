import { readFile } from 'node:fs/promises';
import { loadDocx } from 'ooxml-core/docx';

/** Shared synthetic source for native cached-result input and clipboard comparisons. */
export async function fieldExportSource() {
	const loaded = await loadDocx(
		await readFile(
			new URL('../../core/docx/__fixtures__/field-comments/adjacent-source.docx', import.meta.url),
		),
	);
	const model = {
		...loaded.model,
		blocks: loaded.model.blocks.map((block) => {
			if (block.type !== 'paragraph') return block;
			const identity = block.runs.find((run) => run.field)?.fieldInstanceId;
			return {
				...block,
				runs: block.runs.map((run) =>
					run.field && run.fieldInstanceId === identity ? { ...run, bold: true } : run,
				),
			};
		}),
	};
	return { loaded, model };
}
