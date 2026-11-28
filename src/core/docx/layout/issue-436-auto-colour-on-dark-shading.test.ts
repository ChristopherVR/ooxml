import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

/** The colour each fragment of `Cell <fill>` and `Paragraph <fill>` is drawn in (undefined: black). */
async function colours(name: string): Promise<Map<string, string | undefined>> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	const found = new Map<string, string | undefined>();
	const visit = (blocks: readonly unknown[]) => {
		for (const block of blocks as {
			kind: string;
			lines?: { fragments: { text?: string; color?: string }[] }[];
			rows?: { cells: unknown[][] }[];
		}[]) {
			for (const line of block.lines ?? []) {
				const colours = new Set(line.fragments.map((fragment) => fragment.color));
				if (colours.size !== 1) throw new Error('a line mixes colours');
				const text = line.fragments.map((fragment) => fragment.text ?? '').join('');
				found.set(text.trim(), [...colours][0]);
			}
			for (const row of block.rows ?? []) for (const cell of row.cells) visit(cell);
		}
	};
	for (const page of pages) for (const column of page.columns) visit(column.blocks);
	return found;
}

// Word 16 draws automatic-colour text white on 000000, 371127 and C00000 (cell and paragraph
// shading) and black on 404040, 595959 and 4472C4; the switch is at a luma of 59.5.
it('draws automatic text white on dark cell and paragraph shading (#436)', async () => {
	const found = await colours('issue-436-auto-colour-on-dark-shading');
	for (const fill of ['000000', '371127', 'C00000']) {
		expect(found.get(`Cell ${fill}`), `Cell ${fill}`).toBe('#ffffff');
		expect(found.get(`Paragraph ${fill}`), `Paragraph ${fill}`).toBe('#ffffff');
	}
	for (const fill of ['404040', '595959', '4472C4'])
		expect(found.get(`Cell ${fill}`), `Cell ${fill}`).toBeUndefined();
	expect(found.get('Paragraph 404040')).toBeUndefined();
});

it('keeps an explicit black colour on dark shading (#436 control)', async () => {
	const found = await colours('issue-436-auto-colour-on-dark-shading-control');
	for (const text of ['Cell 000000', 'Cell 371127', 'Cell C00000', 'Paragraph 371127'])
		expect(found.get(text), text).toBe('#000000');
});
