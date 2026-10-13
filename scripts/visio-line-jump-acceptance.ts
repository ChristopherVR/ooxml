// Manual native acceptance input generator. Run through the adjacent PowerShell script.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { editVsdx } from '../src/core/visio/edit';
import { parseVsdx } from '../src/core/visio/parser';
import type { VisioEdit } from '../src/core/visio/edit-commands';

const directory = process.argv[2];
if (!directory) throw new Error('An output directory is required.');
await mkdir(resolve(directory), { recursive: true });
const bytes = new Uint8Array(
	await readFile(
		new URL('../src/core/visio/__fixtures__/native-cached-line-jumps.vsdx', import.meta.url),
	),
);
const pageId = (await parseVsdx(bytes)).pages[0]!.id;
const cases: [string, VisioEdit][] = [
	['off', { type: 'set-page-layout', pageId, lineJumpCode: 0 }],
	['gap', { type: 'format-shape', pageId, shapeId: '1', lineJumpStyle: 2 }],
	['size', { type: 'set-page-layout', pageId, lineJumpFactorX: 2 }],
	['reverse', { type: 'set-page-layout', pageId, lineJumpDirectionX: 2 }],
];
for (const [name, command] of cases) {
	const result = await editVsdx(bytes, [command]);
	if (result.diagnostics.some((note) => note.code === 'edit-line-jump-cache-kept'))
		throw new Error('Native cache was not refreshed.');
	await writeFile(join(directory, `${name}.vsdx`), result.bytes);
}
