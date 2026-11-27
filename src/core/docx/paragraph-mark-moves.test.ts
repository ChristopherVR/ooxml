import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx, type DocumentModel, type Paragraph } from './index';
import { acceptRevision, listRevisions, rejectRevision } from './revision-commands';

const author = 'Ada';
const date = '2026-11-27T00:00:00Z';
const move = { name: 'move1' };
const moved = (id: string, text: string, kind: 'moveFrom' | 'moveTo', n: number): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [{ text, revision: { kind, author, date, id: `${kind}-run-${n}`, move } }],
	markRevision: { kind, author, date, id: `${kind}-mark-${n}`, move },
});
// The Outline view moving "Alpha" (with its body) below "Beta" with Track Changes on.
const model: DocumentModel = {
	...createDocument(),
	blocks: [
		moved('a', 'Alpha', 'moveFrom', 1),
		moved('a-body', 'Alpha body', 'moveFrom', 2),
		{ type: 'paragraph', id: 'b', runs: [{ text: 'Beta' }] },
		moved('a2', 'Alpha', 'moveTo', 1),
		moved('a2-body', 'Alpha body', 'moveTo', 2),
		{ type: 'paragraph', id: 'tail', runs: [{ text: 'Tail' }] },
	],
};
const texts = (value: DocumentModel) =>
	value.blocks.map((block) =>
		block.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : '',
	);

describe('moved paragraph marks', () => {
	it('writes w:moveFrom and w:moveTo paragraph marks inside the move ranges and reads them back', async () => {
		const bytes = await saveDocx(model);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toMatch(/<w:pPr>(?:(?!<\/w:pPr>).)*<w:rPr><w:moveFrom /s);
		expect(xml).toMatch(/<w:pPr>(?:(?!<\/w:pPr>).)*<w:rPr><w:moveTo /s);
		expect(xml).toContain('w:moveFromRangeStart');
		const { model: reopened } = await loadDocx(bytes);
		const marks = reopened.blocks.map((block) =>
			block.type === 'paragraph'
				? [block.markRevision?.kind ?? null, block.markRevision?.move?.name ?? null]
				: null,
		);
		expect(marks).toEqual([
			['moveFrom', 'move1'],
			['moveFrom', 'move1'],
			[null, null],
			['moveTo', 'move1'],
			['moveTo', 'move1'],
			[null, null],
		]);
	});

	it('resolves both sides of the move, paragraph marks included', () => {
		const id = listRevisions(model).find((entry) => entry.kind === 'moveFrom')!.id;
		const accepted = acceptRevision(model, id);
		expect(texts(accepted)).toEqual(['Beta', 'Alpha', 'Alpha body', 'Tail']);
		expect(listRevisions(accepted)).toEqual([]);
		const rejected = rejectRevision(model, id);
		expect(texts(rejected)).toEqual(['Alpha', 'Alpha body', 'Beta', 'Tail']);
		expect(listRevisions(rejected)).toEqual([]);
	});
});
