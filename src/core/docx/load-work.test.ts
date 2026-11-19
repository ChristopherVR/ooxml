// Work budget (docs/performance-budgets.md): how often loading and saving a Word document copies
// its whole model. Each copy costs time and memory in proportion to the document.
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { DocumentModel } from './model';
import { loadDocx } from './parse';
import { docx, p } from './test-support/compare-docs';

afterEach(() => vi.restoreAllMocks());

/** Calls of `structuredClone` on a whole document model while `run` runs. */
async function modelCopies(run: () => Promise<unknown>): Promise<number> {
	const clone = vi.spyOn(globalThis, 'structuredClone');
	try {
		await run();
		return clone.mock.calls.filter(([value]) => Array.isArray((value as DocumentModel)?.blocks))
			.length;
	} finally {
		clone.mockRestore();
	}
}

describe('docx load and save work', () => {
	it('copies the model once to load, and never to save it or an edited copy', async () => {
		const body = Array.from({ length: 50 }, (_, index) => p(`Paragraph ${index}`)).join('');
		const bytes = await docx(body);
		let loaded!: Awaited<ReturnType<typeof loadDocx>>;
		expect(await modelCopies(async () => (loaded = await loadDocx(bytes)))).toBe(1);
		expect(await modelCopies(() => loaded.save())).toBe(0);
		const edited: DocumentModel = { ...loaded.model, blocks: loaded.model.blocks.slice(1) };
		let saved!: Uint8Array;
		expect(await modelCopies(async () => (saved = await loaded.save(edited)))).toBe(0);
		// The work was done: the save wrote the edit.
		const reloaded = await loadDocx(saved);
		expect(reloaded.model.blocks).toHaveLength(49);
	});
});
