import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { saveDocx } from './save.js';
import { acceptRevision } from './revision-commands.js';
import { expectParagraph } from './test-support/access.js';

for (const name of ['alignment', 'spacing', 'indent', 'multiple'])
	describe(`native paragraph ${name} history`, () => {
		const fixture = () =>
			readFile(
				new URL(`./__fixtures__/review-paragraph-formatting/${name}-tracked.docx`, import.meta.url),
			);
		it('preserves the complete prior properties through text editing and both export paths', async () => {
			const loaded = await loadDocx(new Uint8Array(await fixture()));
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			const revision = paragraph.formatRevision!;
			expect(revision.kind).toBe('paragraphChange');
			expect(revision.previousParagraphPropertiesXml).toContain('pPr');
			const model = structuredClone(loaded.model);
			expectParagraph(model.blocks[0]).runs[0]!.text += '!';
			for (const bytes of [await loaded.save(model), await saveDocx(model)]) {
				const actual = expectParagraph((await loadDocx(bytes)).model.blocks[0]);
				expect(actual.formatRevision).toEqual(revision);
				expect(actual.runs[0]!.text).toBe('Paragraph formatting!');
			}
		});
		it('accepts the paragraph formatting revision and exports the current formatting', async () => {
			const loaded = await loadDocx(new Uint8Array(await fixture()));
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			const accepted = acceptRevision(loaded.model, paragraph.formatRevision!.id);
			const bytes = await loaded.save(accepted);
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
			expect(xml).not.toContain('pPrChange');
			const { formatRevision: _revision, ...expected } = paragraph;
			expect(expectParagraph((await loadDocx(bytes)).model.blocks[0])).toEqual(expected);
		});
	});
