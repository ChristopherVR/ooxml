import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from './parse';
import type { Paragraph } from './model';
for (const name of ['begin', 'result', 'whole', 'simple-result', 'adjacent-result'])
	it(`retains native ${name} field comments through a neighboring text edit`, async () => {
		const bytes = new Uint8Array(
			await readFile(new URL(`./__fixtures__/field-comments/${name}.docx`, import.meta.url)),
		);
		const loaded = await loadDocx(bytes);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const ids = loaded.model.comments!.map((comment) => comment.id).sort();
		expect(ids).toHaveLength(2);
		const anchored = paragraph.runs.filter((run) => run.commentIds?.length);
		expect(anchored).toHaveLength(5);
		for (const run of anchored) expect([...run.commentIds!].sort()).toEqual(ids);
		expect(paragraph.runs[0]!.commentIds).toBeUndefined();
		expect(paragraph.runs.at(-1)!.commentIds).toBeUndefined();
		const blocks = [
			{
				...paragraph,
				runs: paragraph.runs.map((run, index) =>
					index === 0 ? { ...run, text: 'Changed ' } : run,
				),
			},
		];
		const saved = await loadDocx(await loaded.save({ ...loaded.model, blocks }));
		expect(
			(saved.model.blocks[0] as Paragraph).runs.filter((run) => run.commentIds?.length),
		).toEqual(anchored);
	});
