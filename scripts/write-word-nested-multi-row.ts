import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { loadDocx } from '../src/core/docx/index';
const directory = process.argv[2];
if (!directory) throw new Error('Provide native probe directory');
const cases = JSON.parse(readFileSync(join(directory, 'native-cases.json'), 'utf8')) as {
	name: string;
	pages: number;
}[];
const records = [];
for (const source of cases) {
	const loaded = await loadDocx(
		new Uint8Array(readFileSync(join(directory, source.name + '.docx'))),
	);
	for (const mode of ['noop', 'edit'] as const) {
		if (mode === 'edit') {
			const table = loaded.model.blocks.find((block) => block.type === 'table');
			if (!table || table.type !== 'table') throw new Error('Missing enclosing table');
			const run = table.rows[0]![0]!.paragraphs[0]!.runs.find((run) => run.text === 'T2R0C0L00');
			if (!run) throw new Error('Missing direct outer text');
			run.text = 'E2R0C0L00';
		}
		const name = source.name + '-core-' + mode;
		writeFileSync(join(directory, name + '.docx'), await loaded.save(loaded.model));
		records.push({
			name,
			pages: source.pages,
			outerText: mode === 'edit' ? 'E2R0C0L00' : 'T2R0C0L00',
		});
	}
}
writeFileSync(join(directory, 'core-export-cases.json'), JSON.stringify(records, null, 2) + '\n');
