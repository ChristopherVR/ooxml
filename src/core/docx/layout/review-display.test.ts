import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph } from '../index';
import { loadDocx } from '../parse';
import { acceptAllRevisions, rejectAllRevisions } from '../revision-commands';
import { adaptDocumentModel } from './adapter';
import { layoutDocumentModel } from './layout';

const measurer = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
const paragraph = (runs: Paragraph['runs']): Paragraph => ({ type: 'paragraph', id: 'p', runs });
const fragments = (result: ReturnType<typeof layoutDocumentModel>) =>
	result.pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph' ? block.lines.flatMap((line) => line.fragments) : [],
			),
		),
	);

describe('paginated review text display', () => {
	for (const [area, name] of [
		['review-formatting', 'multiple'],
		['review-paragraph-formatting', 'multiple'],
	] as const)
		it(`projects native prior ${area} formatting for core layout callers`, async () => {
			const load = async (suffix: string) =>
				loadDocx(
					new Uint8Array(
						await readFile(
							new URL(`../__fixtures__/${area}/${name}-${suffix}.docx`, import.meta.url),
						),
					),
				);
			const before = (await load('before')).model;
			const current = (await load('tracked')).model;
			const source = structuredClone(current);
			const appearance = (result: ReturnType<typeof layoutDocumentModel>) => fragments(result);
			expect(
				appearance(layoutDocumentModel(current, measurer, { reviewDisplayMode: 'original' })),
			).toEqual(appearance(layoutDocumentModel(before, measurer)));
			expect(current).toEqual(source);
		});

	it('hides revisions without shifting later UTF-16 source ranges', () => {
		const model = createDocument();
		model.blocks = [
			paragraph([
				{ text: '😀', revision: { id: 'i', kind: 'insert', author: 'Ada' } },
				{ text: 'Keep ' },
				{ text: 'old', revision: { id: 'd', kind: 'delete', author: 'Ada' } },
				{ text: ' tail' },
			]),
		];
		const before = structuredClone(model);
		for (const [mode, expected] of [
			['original', 'Keep old tail'],
			['final', '😀Keep  tail'],
			['simple', '😀Keep  tail'],
			['all', '😀Keep old tail'],
		] as const) {
			const values = fragments(layoutDocumentModel(model, measurer, { reviewDisplayMode: mode }));
			expect(values.map((value) => value.text).join('')).toBe(expected);
			expect(values.find((value) => value.text === 'Keep')).toMatchObject({
				sourceStart: 2,
				sourceEnd: 6,
			});
			expect(values.find((value) => value.text === 'tail')).toMatchObject({
				sourceStart: 11,
				sourceEnd: 15,
			});
		}
		expect(model).toEqual(before);
	});

	it('suppresses hidden line/page breaks and pictures instead of consuming layout space', () => {
		const model = createDocument();
		model.blocks = [
			paragraph([
				{ text: 'Inserted\n', revision: { id: 'i', kind: 'insert', author: 'Ada' } },
				{ text: '', break: 'page', revision: { id: 'break', kind: 'insert', author: 'Ada' } },
				{
					text: '',
					image: {
						relId: 'r1',
						partName: 'word/media/image.png',
						contentType: 'image/png',
						widthPx: 200,
						heightPx: 200,
					},
					revision: { id: 'image', kind: 'insert', author: 'Ada' },
				},
				{ text: 'Keep' },
			]),
		];
		const result = layoutDocumentModel(model, measurer, { reviewDisplayMode: 'original' });
		expect(result.pages).toHaveLength(1);
		expect(fragments(result)).toMatchObject([{ text: 'Keep', sourceStart: 11, sourceEnd: 15 }]);
	});

	it('numbers only visible note references and omits hidden footnotes and endnotes', () => {
		const model = createDocument();
		model.blocks = [
			paragraph([
				{ text: 'Text' },
				{
					text: '',
					noteReference: { id: '1', kind: 'footnote' },
					revision: { id: 'd', kind: 'delete', author: 'Ada' },
				},
				{ text: '', noteReference: { id: '2', kind: 'footnote' } },
				{
					text: '',
					noteReference: { id: '3', kind: 'endnote' },
					revision: { id: 'endnote', kind: 'delete', author: 'Ada' },
				},
			]),
		];
		model.footnotes = ['1', '2'].map((id) => ({
			id,
			blocks: [
				{
					...paragraph([{ text: '', noteMark: 'footnote' }, { text: `Note ${id}` }]),
					id: `note-${id}`,
				},
			],
		}));
		model.endnotes = [
			{ id: '3', blocks: [{ ...paragraph([{ text: 'Endnote' }]), id: 'endnote-3' }] },
		];
		const input = adaptDocumentModel(model, undefined, { reviewDisplayMode: 'final' });
		const body = input.sections[0]!.blocks[0]!;
		if (body.kind !== 'paragraph') throw new Error('Expected paragraph');
		expect(body.footnotes?.map((note) => note.id)).toEqual(['2']);
		expect(body.footnotes![0]!.paragraphs[0]!.runs[0]!.text).toBe('1');
		expect(input.sections[0]!.blocks.map((block) => block.id)).not.toContain('endnote-3');
	});

	it('matches resolved text for the native Word move reference while retaining source revisions', async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(new URL('../__fixtures__/review-moves/word-saved.docx', import.meta.url)),
			),
		);
		const before = structuredClone(loaded.model);
		for (const mode of ['original', 'final'] as const) {
			const expected =
				mode === 'original' ? rejectAllRevisions(loaded.model) : acceptAllRevisions(loaded.model);
			expect(
				fragments(layoutDocumentModel(loaded.model, measurer, { reviewDisplayMode: mode }))
					.map((run) => run.text)
					.join(''),
			).toBe(
				fragments(layoutDocumentModel(expected, measurer))
					.map((run) => run.text)
					.join(''),
			);
		}
		expect(loaded.model).toEqual(before);
	});
});
