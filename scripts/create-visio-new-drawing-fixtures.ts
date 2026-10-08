import { mkdir, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { createVsdx } from '../src/core/visio/create-document';
import { editVsdx } from '../src/core/visio/edit';

const directory = process.argv[2];
if (!directory) throw new Error('Provide an output directory for owned native test files.');
const output = resolve(directory);
await mkdir(output, { recursive: true });
const cases = [
	{ name: 'letter', width: 8.5, height: 11 },
	{ name: 'landscape', width: 11, height: 8.5 },
	{ name: 'a4', width: 210 / 25.4, height: 297 / 25.4 },
];
for (const item of cases) {
	const blank = await createVsdx(item);
	await writeFile(join(output, `${item.name}-blank.vsdx`), blank);
	const drawn = await editVsdx(blank, [
		{
			type: 'create-rectangle',
			pageId: '0',
			shapeId: '1',
			x: 2,
			y: 3,
			width: 2,
			height: 1,
			text: 'New diagram',
		},
	]);
	await writeFile(join(output, `${item.name}-drawn.vsdx`), drawn.bytes);
}
await writeFile(join(output, 'cases.json'), JSON.stringify(cases, null, 2));
console.log(output);
