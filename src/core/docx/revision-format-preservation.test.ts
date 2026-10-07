import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { saveDocx } from './save.js';
import { acceptRevision, rejectRevision, rejectAllRevisions } from './revision-commands.js';
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
		it('restores the properties recorded by native Word rejection', async () => {
			const loaded = await loadDocx(new Uint8Array(await fixture(name)));
			const native = await loadDocx(
				new Uint8Array(
					await readFile(
						new URL(`./__fixtures__/review-formatting/${name}-rejected.docx`, import.meta.url),
					),
				),
			);
			const revision = expectParagraph(loaded.model.blocks[0]).runs.find(
				(run) => run.revision,
			)!.revision!;
			for (const model of [
				rejectRevision(loaded.model, revision.id),
				rejectAllRevisions(loaded.model),
			]) {
				const restored = expectParagraph(model.blocks[0]);
				expect(restored.runs.map(({ restoredRunPropertiesXml: _snapshot, ...run }) => run)).toEqual(
					expectParagraph(native.model.blocks[0]).runs,
				);
				for (const bytes of [await loaded.save(model), await saveDocx(model)]) {
					const reopened = await loadDocx(bytes);
					expect(expectParagraph(reopened.model.blocks[0]).runs).toEqual(
						expectParagraph(native.model.blocks[0]).runs,
					);
					const xml = await (
						await JSZip.loadAsync(bytes)
					)
						.file('word/document.xml')!
						.async('string');
					expect(xml).not.toContain('rPrChange');
					// Complex-script fonts survive the full prior-properties restoration.
					expect(xml).toContain('w:cs="Arial"');
				}
			}
		});
		it('retains the restored XML basis while later edits overlay known properties', async () => {
			const loaded = await loadDocx(new Uint8Array(await fixture(name)));
			const revision = expectParagraph(loaded.model.blocks[0]).runs.find(
				(run) => run.revision,
			)!.revision!;
			const model = rejectRevision(loaded.model, revision.id);
			const run = expectParagraph(model.blocks[0]).runs.find(
				(run) => run.restoredRunPropertiesXml,
			)!;
			run.text += '!';
			run.fontSize = 18;
			const bytes = await loaded.save(model);
			const reopened = await loadDocx(bytes);
			const actual = expectParagraph(reopened.model.blocks[0]).runs.find((run) =>
				run.text.includes('Format'),
			)!;
			expect(actual).toMatchObject({ text: 'Format me!', fontSize: 18 });
			expect(actual.revision).toBeUndefined();
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
			expect(xml).toContain('w:cs="Arial"');
		});
	});
