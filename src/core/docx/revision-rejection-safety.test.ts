import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph, type Revision } from './model.js';
import { acceptAllRevisions, rejectAllRevisions } from './revision-commands.js';

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('atomic rejection failures', () => {
	it.each(['formatChange', 'paragraphChange'] as const)(
		'preserves %s revisions with missing or malformed prior properties',
		(kind) => {
			for (const snapshot of [undefined, '<wrong/>', '<!DOCTYPE x><x/>']) {
				const model = createDocument();
				const revision: Revision = { id: 'format', kind, author: 'Ada' };
				if (snapshot !== undefined) {
					if (kind === 'formatChange') revision.previousRunPropertiesXml = snapshot;
					else revision.previousParagraphPropertiesXml = snapshot;
				}
				const paragraph: Paragraph = {
					type: 'paragraph',
					id: 'p',
					align: 'center',
					runs: [
						{ text: 'Text', bold: true, ...(kind === 'formatChange' ? { revision } : {}) },
						{ text: 'Pending', revision: { id: 'insert', kind: 'insert', author: 'Grace' } },
					],
					...(kind === 'paragraphChange' ? { formatRevision: revision } : {}),
				};
				model.blocks = [paragraph];
				const before = structuredClone(model);
				expect(() => rejectAllRevisions(model)).toThrow();
				expect(model).toEqual(before);
				expect(acceptAllRevisions(model).blocks).toEqual([
					{
						type: 'paragraph',
						id: 'p',
						align: 'center',
						runs: [{ text: 'Text', bold: true }, { text: 'Pending' }],
					},
				]);
			}
		},
	);

	it('preserves a paragraph mark when its merge cannot be completed', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'last',
				runs: [{ text: 'Last' }],
				markRevision: { id: 'mark', kind: 'insert', author: 'Ada' },
			},
		];
		const before = structuredClone(model);
		expect(() => rejectAllRevisions(model)).toThrow('no following paragraph');
		expect(model).toEqual(before);
	});

	it('rejects valid formatting snapshots and text insertions together', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				align: 'center',
				formatRevision: {
					id: 'p-format',
					kind: 'paragraphChange',
					author: 'Ada',
					previousParagraphPropertiesXml: `<w:pPr xmlns:w="${wordNamespace}"/>`,
				},
				runs: [
					{
						text: 'Text',
						bold: true,
						revision: {
							id: 'r-format',
							kind: 'formatChange',
							author: 'Ada',
							previousRunPropertiesXml: `<w:rPr xmlns:w="${wordNamespace}"/>`,
						},
					},
					{ text: 'Pending', revision: { id: 'insert', kind: 'insert', author: 'Grace' } },
				],
			},
		];
		const result = rejectAllRevisions(model);
		expect(result.blocks[0]).toMatchObject({ runs: [{ text: 'Text' }] });
		const paragraph = result.blocks[0] as Paragraph;
		expect(paragraph.align).toBeUndefined();
		expect(paragraph.formatRevision).toBeUndefined();
		expect(paragraph.runs).toHaveLength(1);
		expect(paragraph.runs[0]?.bold).toBeUndefined();
		expect(paragraph.runs[0]?.revision).toBeUndefined();
	});
});
