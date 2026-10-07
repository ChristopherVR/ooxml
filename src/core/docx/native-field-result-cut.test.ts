import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from './parse';
import type { Paragraph } from './model';
for (const kind of ['result', 'partial'])
	it(`preserves native ${kind} field result move and copied formatting on export`, async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(new URL(`./__fixtures__/field-result-cut/${kind}.docx`, import.meta.url)),
			),
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const results: string[] = [];
		let text = '';
		for (const run of paragraph.runs) {
			if (run.field) text += run.text;
			if (run.fieldChar === 'end') {
				results.push(text);
				text = '';
			}
		}
		expect(results).toEqual([kind === 'result' ? '' : 'ACDE', 'ABCDE']);
		expect(
			paragraph.runs.find((run) => !run.field && run.text === (kind === 'result' ? 'ABCDE' : 'B')),
		).toMatchObject({ bold: true });
		const runs = paragraph.runs.map((run, index) =>
			index === 0 ? { ...run, text: 'Changed ' } : run,
		);
		const reloaded = await loadDocx(
			await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] }),
		);
		expect((reloaded.model.blocks[0] as Paragraph).runs.slice(1)).toEqual(paragraph.runs.slice(1));
	});
