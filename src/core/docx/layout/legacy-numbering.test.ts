import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx, saveDocx } from '../index';
import type { Paragraph } from '../model';
import { createFakeMeasurer, layoutDocumentModel } from './index';
import { at } from './test-helpers';

const load = async () =>
	(
		await loadDocx(
			await readFile(
				new URL('./fixtures/word-parity/issue-371-legacy-indent.docx', import.meta.url),
			),
		)
	).model;

// Issue #371: a `w:legacy` level starts the text after the bullet at `w:legacyIndent` (283
// twips), not at the next default tab stop (720 twips).
it('puts the text after a legacy bullet at the legacy indent', async () => {
	const { pages } = layoutDocumentModel(await load(), createFakeMeasurer());
	const lines = at(at(pages, 0).columns, 0).blocks.flatMap((block) =>
		block.kind === 'paragraph' ? block.lines : [],
	);
	expect(lines).toHaveLength(2);
	const text = at(lines, 0).fragments.find((fragment) => fragment.text.startsWith('First'));
	expect(text?.xPx).toBeCloseTo((283 / 20) * (96 / 72), 1);
});

const textX = (model: Awaited<ReturnType<typeof load>>) => {
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	const line = at(
		at(at(pages, 0).columns, 0).blocks.flatMap((block) =>
			block.kind === 'paragraph' ? block.lines : [],
		),
		0,
	);
	return line.fragments.find((fragment) => fragment.text.startsWith('First'))?.xPx;
};

it('measures the legacy indent from the number start, not the margin', async () => {
	const model = await load();
	Object.assign(model.blocks[0] as Paragraph, { indentLeftTwips: 1440 });
	expect(textX(model)).toBeCloseTo(96 + (283 / 20) * (96 / 72), 1);
});

it('leaves at least legacySpace between the number and the text', async () => {
	const model = await load();
	const level = at(Object.values(model.numberingCatalog?.abstractNums['0']?.levels ?? {}), 0);
	level.legacy = { enabled: true, indentTwips: level.legacy?.indentTwips ?? 0, spaceTwips: 600 } as never;
	// The bullet's own width plus 600 twips (40px) is further than the 283 twip indent.
	expect(textX(model)).toBeGreaterThan(40);
});

it('keeps w:legacy when the document is saved, edited or not', async () => {
	const model = await load();
	const level = model.numberingCatalog?.abstractNums['0']?.levels[0];
	expect(level?.legacy).toEqual({ enabled: true, indentTwips: 283, spaceTwips: 0 });
	const check = async (bytes: Uint8Array) => {
		const zip = await JSZip.loadAsync(bytes);
		const xml = (await zip.file('word/numbering.xml')?.async('string')) ?? '';
		expect(xml).toContain('w:legacyIndent="283"');
		expect(xml).toContain('w:legacySpace="0"');
	};
	await check(await saveDocx(structuredClone(model)));
	const edited = structuredClone(model);
	const paragraph = edited.blocks[0] as Paragraph;
	at(paragraph.runs, 0).text = 'Edited bulleted paragraph.';
	await check(await saveDocx(edited));
});
