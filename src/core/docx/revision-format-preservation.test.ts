import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { saveDocx } from './save.js';
import { acceptRevision } from './revision-commands.js';
import { expectParagraph } from './test-support/access.js';

const fixture = (name: string) =>
	readFile(new URL(`./__fixtures__/review-formatting/${name}-tracked.docx`, import.meta.url));
for (const name of ['bold', 'multiple'])
	describe(`native ${name} formatting snapshot`, () => {
		it('retains historical properties through text editing and standalone export', async () => {
			const loaded = await loadDocx(new Uint8Array(await fixture(name)));
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			const run = paragraph.runs.find((item) => item.revision?.kind === 'formatChange')!;
			expect(run.revision!.previousRunPropertiesXml).toContain('rPr');
			const updated = {
				...loaded.model,
				blocks: [
					{
						...paragraph,
						runs: paragraph.runs.map((item) =>
							item === run ? { ...item, text: `${item.text}!` } : item,
						),
					},
				],
			};
			for (const bytes of [await loaded.save(updated), await saveDocx(updated)]) {
				const reopened = await loadDocx(bytes);
				const actual = expectParagraph(reopened.model.blocks[0]).runs.find(
					(item) => item.revision?.kind === 'formatChange',
				)!;
				expect(actual.text).toBe(`${run.text}!`);
				expect(actual.revision).toEqual(run.revision);
				const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
				expect(xml).toContain('w:rPrChange');
			}
		});
		it('clears the historical snapshot when accepting the formatting revision', async () => {
			const loaded = await loadDocx(new Uint8Array(await fixture(name)));
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			const revision = paragraph.runs.find(
				(run) => run.revision?.kind === 'formatChange',
			)!.revision!;
			const accepted = acceptRevision(loaded.model, revision.id);
			const bytes = await loaded.save(accepted);
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
			expect(xml).not.toContain('rPrChange');
		});
	});
