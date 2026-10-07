import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from './parse';
import type { Paragraph } from './model';

for (const [kind, text] of Object.entries({
	partial: 'AXCDE',
	whole: 'X',
	start: 'XBCDE',
	end: 'ABCDX',
	empty: '',
	'tracked-empty': 'ABCDE',
}))
	it(`preserves native ${kind} cached field replacement through a neighboring edit`, async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(
					new URL(`./__fixtures__/field-result-replace/${kind}.docx`, import.meta.url),
				),
			),
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const fields = paragraph.runs.filter((run) => run.field);
		if (kind === 'tracked-empty') expect(fields[0]!.revision?.kind).toBe('delete');
		// Native Word converts imported simple fields to complex fields when saving.
		const results: string[] = [];
		let result = '';
		for (const run of paragraph.runs) {
			if (run.field) result += run.text;
			if (run.fieldChar === 'end') {
				results.push(result);
				result = '';
			}
		}
		expect(results).toEqual([text, 'ABCDE']);
		for (const run of fields.slice(0, -1)) {
			expect(run.field!.instr.trim()).toBe('QUOTE "ABCDE"');
			expect(run.bold).toBe(true);
		}
		const runs = paragraph.runs.map((run, index) =>
			index === 0 ? { ...run, text: 'Changed ' } : run,
		);
		const reloaded = await loadDocx(
			await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] }),
		);
		expect((reloaded.model.blocks[0] as Paragraph).runs.slice(1)).toEqual(paragraph.runs.slice(1));
	});
