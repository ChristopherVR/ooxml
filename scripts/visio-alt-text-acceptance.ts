import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx } from '../src/core/visio/edit';
import { parseVsdx } from '../src/core/visio/parser';

const directory = process.argv[2];
if (!directory) throw new Error('An output directory is required.');
const bytes = new Uint8Array(await readFile(join(directory, 'native.vsdx')));
const page = (await parseVsdx(bytes)).pages[0]!;
const shape = page.shapes[0]!;
if (
	page.accessibilityTitle !== 'Native page title' ||
	page.alternativeText !== 'Native page description'
)
	throw new Error('Native page metadata was not parsed.');
if (shape.accessibilityTitle !== 'Native title' || shape.alternativeText !== 'Native description')
	throw new Error('Native metadata was not parsed.');
const result = await editVsdx(bytes, [
	{
		type: 'set-shape-alt-text',
		pageId: page.id,
		shapeId: shape.id,
		title: 'Service status',
		description: 'Green rectangle\nAll services operational.',
	},
]);
const pageResult = await editVsdx(result.bytes, [
	{
		type: 'set-page-alt-text',
		pageId: page.id,
		title: 'Process overview',
		description: 'Requests arrive\nthen proceed to approval.',
	},
]);
await writeFile(join(directory, 'edited.vsdx'), pageResult.bytes);
const cleared = await editVsdx(pageResult.bytes, [
	{ type: 'set-shape-alt-text', pageId: page.id, shapeId: shape.id, title: '', description: '' },
]);
const pageCleared = await editVsdx(cleared.bytes, [
	{ type: 'set-page-alt-text', pageId: page.id, title: '', description: '' },
]);
await writeFile(join(directory, 'cleared.vsdx'), pageCleared.bytes);
