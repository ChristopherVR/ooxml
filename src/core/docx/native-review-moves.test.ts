import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { acceptRevision, listRevisions, rejectRevision } from './revision-commands';

const fixture = (name: string) =>
	readFile(new URL(`./__fixtures__/review-moves/${name}`, import.meta.url));
const entries = (model: Parameters<typeof listRevisions>[0]) =>
	listRevisions(model).map(({ kind, author, move }) => ({ kind, author, name: move?.name }));

describe('native Word move export references', () => {
	it('loads the same two linked moves before and after Word saves the core export', async () => {
		const source = await loadDocx(new Uint8Array(await fixture('core-created.docx')));
		const native = await loadDocx(new Uint8Array(await fixture('word-saved.docx')));
		expect(entries(source.model)).toEqual(entries(native.model));
		expect(entries(native.model)).toEqual([
			{ kind: 'moveFrom', author: 'Ada', name: 'move1' },
			{ kind: 'moveFrom', author: 'Ada', name: 'move2' },
			{ kind: 'moveTo', author: 'Ada', name: 'move1' },
			{ kind: 'moveTo', author: 'Ada', name: 'move2' },
		]);
		const xml = await (
			await JSZip.loadAsync(await fixture('core-created.docx'))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).not.toContain('w:delText');
	});
	it('matches native acceptance of one move without resolving the other', async () => {
		const source = await loadDocx(new Uint8Array(await fixture('word-saved.docx')));
		const native = await loadDocx(new Uint8Array(await fixture('word-accepted-first.docx')));
		const revision = listRevisions(source.model).find(
			(entry) => entry.kind === 'moveFrom' && entry.move?.name === 'move1',
		)!;
		const accepted = acceptRevision(source.model, revision.id);
		expect(entries(accepted)).toEqual(entries(native.model));
		expect(
			accepted.blocks.map((block) =>
				block.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : '',
			),
		).toEqual(
			native.model.blocks.map((block) =>
				block.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : '',
			),
		);
		const reopened = await loadDocx(await source.save(accepted));
		expect(entries(reopened.model)).toEqual(entries(native.model));
	});
	it('keeps moved-from text when rejecting the first move', async () => {
		const source = await loadDocx(new Uint8Array(await fixture('word-saved.docx')));
		const native = await loadDocx(new Uint8Array(await fixture('word-rejected-first.docx')));
		const revision = listRevisions(source.model).find((entry) => entry.move?.name === 'move1')!;
		const rejected = rejectRevision(source.model, revision.id);
		expect(entries(rejected).map((entry) => entry.name)).toEqual(['move2', 'move2']);
		const first = rejected.blocks[0]!;
		if (first.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(first.runs.map((run) => run.text).join('')).toBe('FirstMove');
		expect(entries(rejected)).toEqual(entries(native.model));
		expect(
			rejected.blocks.map((block) =>
				block.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : '',
			),
		).toEqual(
			native.model.blocks.map((block) =>
				block.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : '',
			),
		);
		const reopened = await loadDocx(await source.save(rejected));
		expect(entries(reopened.model)).toEqual(entries(native.model));
	});
});
