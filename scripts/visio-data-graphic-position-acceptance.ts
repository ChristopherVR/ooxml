import { writeFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import { createVsdx } from '../src/core/visio/create-document';
import { editVsdx } from '../src/core/visio/edit';
import { parseVsdx } from '../src/core/visio/parser';
import { visioDataGraphicEdits, visioDataGraphicParts } from '../src/core/visio/ui/data-graphics';
const directory = process.argv[2];
if (!directory) throw new Error('Output directory required.');
await mkdir(directory, { recursive: true });
let bytes = (
	await editVsdx(await createVsdx(), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 3, width: 2, height: 1 },
		{
			type: 'set-shape-data',
			pageId: '0',
			shapeId: '1',
			row: 'Load',
			data: { label: 'Load', prompt: '', type: 'number', format: '', value: '42' },
		},
	])
).bytes;
for (const kind of ['text', 'bar', 'icon'] as const) {
	const page = (await parseVsdx(bytes)).pages[0]!;
	bytes = (
		await editVsdx(
			bytes,
			visioDataGraphicEdits(page, ['1'], { kind, field: 'Load', min: 0, max: 100 }),
		)
	).bytes;
}
await writeFile(join(directory, 'authored.vsdx'), bytes);
const moved = await editVsdx(bytes, [
	{ type: 'move-shape', pageId: '0', shapeId: '1', x: 4, y: 5 },
]);
await writeFile(join(directory, 'core-moved.vsdx'), moved.bytes);
const page = (await parseVsdx(bytes)).pages[0]!;
const parts = new Set(
	visioDataGraphicParts(page)
		.get('1')!
		.map((part) => part.id),
);
await writeFile(
	join(directory, 'expected.json'),
	JSON.stringify(
		page.shapes
			.filter((shape) => parts.has(shape.id))
			.map((shape) => ({ id: Number(shape.id), x: shape.rotation!.pinX, y: shape.rotation!.pinY })),
		null,
		2,
	),
);
