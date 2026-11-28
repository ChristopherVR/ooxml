import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx, type DocumentModel, type Paragraph } from './index';
import { acceptRevision, rejectRevision } from './revision-commands';
import JSZip from 'jszip';

const ins = { kind: 'insert' as const, author: 'Ada', date: '2026-01-01T00:00:00Z', id: '4' };
const del = { kind: 'delete' as const, author: 'Bob', date: '2026-02-02T00:00:00Z', id: '5' };
const model: DocumentModel = {
	...createDocument(),
	blocks: [
		{ type: 'paragraph', id: 'a', runs: [{ text: 'Alpha ' }], markRevision: { ...del, id: '6' } },
		{
			type: 'paragraph',
			id: 'b',
			runs: [{ text: '' }],
			markRevision: del,
			markInsertion: ins,
		},
		{ type: 'paragraph', id: 'c', runs: [{ text: 'tail' }] },
	],
};
const texts = (value: DocumentModel) =>
	value.blocks.map((block) =>
		block.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : '',
	);

describe('a paragraph mark that was inserted and then deleted', () => {
	it('is written with w:ins then w:del and read back with both', async () => {
		const bytes = await saveDocx(model);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toMatch(
			/<w:rPr><w:ins [^>]*w:id="4"[^>]*w:author="Ada"[^>]*\/><w:del [^>]*w:id="5"[^>]*w:author="Bob"[^>]*\/><\/w:rPr>/,
		);
		const { model: reopened } = await loadDocx(bytes);
		const mark = reopened.blocks[1] as Paragraph;
		expect(mark.markRevision).toMatchObject({ kind: 'delete', id: '5', author: 'Bob' });
		expect(mark.markRevision?.date).toBe(del.date);
		expect(mark.markInsertion).toMatchObject({ kind: 'insert', id: '4', author: 'Ada' });
		expect(mark.markInsertion?.date).toBe(ins.date);
		// Saving again keeps both.
		const again = await loadDocx(await saveDocx(reopened));
		expect((again.model.blocks[1] as Paragraph).markInsertion?.id).toBe('4');
	});

	it('accepting the deletion joins the paragraphs; rejecting it leaves an inserted mark', () => {
		expect(texts(acceptRevision(model, '5'))).toEqual(['Alpha ', 'tail']);
		const rejected = rejectRevision(model, '5');
		const mark = rejected.blocks.find((block) => block.id === 'b') as Paragraph;
		expect(mark.markRevision).toEqual(ins);
		expect(mark.markInsertion).toBeUndefined();
	});
});
