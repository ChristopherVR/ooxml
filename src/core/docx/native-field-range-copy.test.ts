import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from './parse';
import type { Paragraph } from './model';
for (const kind of ['whole', 'result', 'partial'])
	it(`retains native ${kind} field range transfer through a neighboring text edit`, async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(new URL(`./__fixtures__/field-range-copy/${kind}.docx`, import.meta.url)),
			),
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const fields = paragraph.runs.filter((run) => run.field);
		expect(fields).toHaveLength(kind === 'whole' ? 3 : 2);
		if (kind !== 'whole')
			expect(
				paragraph.runs.find(
					(run) => !run.field && run.text === (kind === 'result' ? 'ABCDE' : 'B'),
				),
			).toMatchObject({ bold: true });
		const runs = paragraph.runs.map((run, index) =>
			index === 0 ? { ...run, text: 'Changed ' } : run,
		);
		const reloaded = await loadDocx(
			await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] }),
		);
		expect((reloaded.model.blocks[0] as Paragraph).runs.slice(1)).toEqual(paragraph.runs.slice(1));
	});
