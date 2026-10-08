import { mkdir, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { createVsdx } from '../src/core/visio/create-document';
import { editVsdx } from '../src/core/visio/edit';
import { openEditablePackage, writeEditedPackage } from '../src/core/visio/edit-package';
import { DEFAULTS } from '../src/core/visio/package-common';

const directory = process.argv[2];
if (!directory) throw new Error('Provide a directory for owned native test files.');
const output = resolve(directory);
await mkdir(output, { recursive: true });
const cases = [];
const values = { x: 2.75, y: 2.75, width: 2.5, height: 1.25, angle: -35 };
for (const scale of [0.5, 1, 2]) {
	const { parts } = await openEditablePackage(await createVsdx(), DEFAULTS, () => {});
	const path = 'visio/pages/pages.xml';
	parts.set(
		path,
		new TextEncoder().encode(
			new TextDecoder()
				.decode(parts.get(path)!)
				.replace('N="DrawingScale" V="1"', `N="DrawingScale" V="${1 / scale}"`),
		),
	);
	const scaled = await writeEditedPackage(
		parts,
		DEFAULTS.maxInputBytes,
		Date.now() + DEFAULTS.maxRuntimeMs,
		() => {},
	);
	const initial = await editVsdx(scaled, [
		{
			type: 'create-rectangle',
			pageId: '0',
			shapeId: '1',
			x: 2,
			y: 3,
			width: 2,
			height: 1,
			text: 'Numeric shape',
		},
	]);
	const rotated = await editVsdx(initial.bytes, [
		{ type: 'rotate-shape', pageId: '0', shapeId: '1', angle: Math.PI / 6 },
	]);
	for (const [field, value] of Object.entries(values)) {
		const name = `${scale}-${field}`;
		await writeFile(join(output, `${name}-original.vsdx`), rotated.bytes);
		cases.push({ name, scale, field, value });
	}
}
await writeFile(join(output, 'cases.json'), JSON.stringify(cases, null, 2));
console.log(output);
