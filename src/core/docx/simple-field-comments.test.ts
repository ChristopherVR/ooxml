import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import type { Paragraph } from './model';

it('retains imported partial result scopes inside a field while enclosing whole-field scopes', async () => {
	const loaded = await loadDocx(
		new Uint8Array(
			await readFile(new URL('./__fixtures__/field-comments/simple-source.docx', import.meta.url)),
		),
	);
	const paragraph = loaded.model.blocks[0] as Paragraph;
	const result = paragraph.runs[1]!;
	const runs = [
		paragraph.runs[0]!,
		{ ...result, text: 'AB', commentIds: ['0'] },
		{ ...result, text: 'CDE', commentIds: ['0', '1'], bold: true },
		paragraph.runs[2]!,
	];
	const bytes = await loaded.save({
		...loaded.model,
		blocks: [{ ...paragraph, runs }],
		comments: ['0', '1'].map((id) => ({ id, author: id, text: id })),
	});
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	expect(xml.indexOf('<w:commentRangeStart w:id="0"')).toBeLessThan(xml.indexOf('<w:fldSimple'));
	expect(xml.indexOf('<w:commentRangeStart w:id="1"')).toBeGreaterThan(xml.indexOf('<w:fldSimple'));
	expect(xml.indexOf('<w:commentRangeEnd w:id="1"')).toBeLessThan(xml.indexOf('</w:fldSimple>'));
	expect(xml.indexOf('<w:commentRangeEnd w:id="0"')).toBeGreaterThan(xml.indexOf('</w:fldSimple>'));
	const reloaded = await loadDocx(bytes);
	expect(
		(reloaded.model.blocks[0] as Paragraph).runs
			.filter((run) => run.field)
			.map((run) => run.commentIds?.sort()),
	).toEqual([['0'], ['0', '1']]);
});

it('exports whole simple-field comments outside the wrapper and retains overlapping scopes', async () => {
	const loaded = await loadDocx(
		new Uint8Array(
			await readFile(new URL('./__fixtures__/field-comments/simple-source.docx', import.meta.url)),
		),
	);
	const paragraph = loaded.model.blocks[0] as Paragraph;
	const runs = paragraph.runs.map((run) => (run.field ? { ...run, commentIds: ['0', '1'] } : run));
	const bytes = await loaded.save({
		...loaded.model,
		blocks: [{ ...paragraph, runs }],
		comments: ['0', '1'].map((id) => ({ id, author: id, text: id })),
	});
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(
		/<w:commentRangeStart w:id="0"\/><w:commentRangeStart w:id="1"\/><w:fldSimple[^>]*><w:r><w:t>ABCDE<\/w:t><\/w:r><\/w:fldSimple><w:commentRangeEnd w:id="1"\/>/,
	);
	const reloaded = await loadDocx(bytes);
	expect((reloaded.model.blocks[0] as Paragraph).runs[1]!.commentIds?.sort()).toEqual(['0', '1']);
	expect((reloaded.model.blocks[0] as Paragraph).runs[0]!.commentIds).toBeUndefined();
});
