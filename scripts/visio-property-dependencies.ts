import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx } from '../src/core/visio/edit';
import { parseVsdx } from '../src/core/visio/parser';
const directory = process.argv[2];
if (!directory) throw new Error('Output directory required.');
const bytes = new Uint8Array(await readFile(join(directory, 'native.vsdx')));
const page = (await parseVsdx(bytes)).pages[0]!;
const result = await editVsdx(bytes, [
	{ type: 'resize-shape', pageId: page.id, shapeId: '1', width: 8, height: 2 },
]);
const saved = (await parseVsdx(result.bytes)).pages[0]!;
if (saved.shapes[0]?.shapeData?.find((row) => row.name === 'Offset')?.value !== 2)
	throw new Error('Core did not update the property cache.');
await writeFile(join(directory, 'edited.vsdx'), result.bytes);
