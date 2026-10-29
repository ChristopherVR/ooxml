// Builds a drawing in which a text box follows a rectangle (PinX/PinY read Sheet.1), deletes the
// rectangle with the core API and writes after.vsdx plus expected.json (the follower's PinX and PinY
// caches our XML claims) for scripts/record-visio-delete-freeze.ps1.
// Run: bun scripts/create-visio-delete-freeze-fixture.ts <OutputDirectory>
import { mkdir, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { createVsdx } from '../src/core/visio/create-document';
import { editVsdx } from '../src/core/visio/edit';
import { parseVsdx } from '../src/core/visio/index';

const directory = process.argv[2];
if (!directory) throw new Error('Provide a directory for owned native test files.');
const output = resolve(directory);
await mkdir(output, { recursive: true });
const created = await editVsdx(await createVsdx(), [
	{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 2, width: 1, height: 1 },
	{
		type: 'create-text-box',
		pageId: '0',
		shapeId: '2',
		x: 3,
		y: 1.2,
		width: 1,
		height: 0.2,
		text: 'Metric',
		followShapeId: '1',
	},
]);
const moved = await editVsdx(created.bytes, [
	{ type: 'move-shape', pageId: '0', shapeId: '1', x: 4, y: 5 },
]);
const deleted = await editVsdx(moved.bytes, [{ type: 'delete-shape', pageId: '0', shapeId: '1' }]);
await writeFile(join(output, 'before.vsdx'), moved.bytes);
await writeFile(join(output, 'after.vsdx'), deleted.bytes);
const follower = (await parseVsdx(deleted.bytes)).pages[0]!.shapes.find((s) => s.id === '2')!;
await writeFile(
	join(output, 'expected.json'),
	JSON.stringify({ shape: '2', pinX: follower.rotation!.pinX, pinY: follower.rotation!.pinY }),
);
console.log(output);
