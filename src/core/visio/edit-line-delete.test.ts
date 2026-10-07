import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, row, section, shape } from './test-fixtures';

const remove = { type: 'delete-shape' as const, pageId: '0', shapeId: '1' };
const line = (extra = '', geometry = '') =>
	shape(
		'1',
		cell('Width', 2) +
			cell('Height', 0) +
			cell('BeginX', 1) +
			cell('BeginY', 1.5) +
			cell('EndX', 3) +
			cell('EndY', 1.5) +
			extra +
			geometry,
	);
const control = shape('2', cell('Width', 1) + cell('Height', 1) + rectangle);
const source = (target = line(), trailing = '', document = '', other = control) =>
	fixture({
		document,
		pages: [{ id: '0', contents: `<Shapes>${target}${other}</Shapes>${trailing}` }],
		edit: (zip) => zip.file('unknown/untouched.bin', new Uint8Array([1, 3, 5, 255])),
	});

it.each([
	line(),
	line(cell('LockBegin', 1) + cell('LockEnd', 1)).replace(
		'N="BeginX" V="1"',
		'N="BeginX" V="1" F="GUARD(1)"',
	),
	line(
		'',
		section(
			'Geometry',
			row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) + row(2, 'ArcTo', cell('X', 2) + cell('Y', 0)),
		),
	),
	shape('1', cell('BeginX', 1) + cell('EndX', 1) + cell('Height', 0) + cell('Width', 0)),
])(
	'deletes an unreferenced local line without rewriting its transform or geometry',
	async (target) => {
		const bytes = await source(target);
		const result = await editVsdx(bytes, [remove]);
		const parsed = await parseVsdx(result.bytes);
		expect(parsed.pages[0]!.shapes.map((shape) => shape.id)).toEqual(['2']);
		expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
		const original = await JSZip.loadAsync(bytes),
			saved = await JSZip.loadAsync(result.bytes);
		for (const [name, entry] of Object.entries(original.files))
			if (!entry.dir && name !== 'visio/pages/page1.xml')
				expect(await saved.file(name)!.async('uint8array'), name).toEqual(
					await entry.async('uint8array'),
				);
	},
);

it('checks local and inherited LockDelete, references and Connects with the shared delete guard', async () => {
	await expect(editVsdx(await source(line(cell('LockDelete', 1))), [remove])).rejects.toThrow(
		'LockDelete',
	);
	await expect(
		editVsdx(
			await source(
				line(),
				'',
				`<StyleSheets><StyleSheet ID="0">${cell('LockDelete', 1)}</StyleSheet></StyleSheets>`,
			),
			[remove],
		),
	).rejects.toThrow('Inherited protection');
	const dependent = shape(
		'2',
		cell('PinX', 3, 'Sheet.1!EndX') + cell('Width', 1) + cell('Height', 1) + rectangle,
	);
	const bytes = await source(line(), '', '', dependent),
		original = bytes.slice();
	await expect(editVsdx(bytes, [remove])).rejects.toThrow('referenced by a ShapeSheet formula');
	expect(bytes).toEqual(original);
	await expect(
		editVsdx(
			await source(
				line(),
				'<Connects><Connect FromSheet="1" ToSheet="2" FromCell="BeginX" ToCell="PinX"/></Connects>',
			),
			[remove],
		),
	).rejects.toThrow('Connect record');
});

const directory = process.env.VISIO_NATIVE_LINE_DELETION_DIR;
it.skipIf(!directory)(
	'matches real native line deletion and preserves the surviving native shape',
	async () => {
		const bytes = await readFile(join(directory!, 'original.vsdx'));
		const native = await parseVsdx(await readFile(join(directory!, 'deleted.vsdx')));
		const original = await parseVsdx(bytes);
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			deletedShapeIds: string[];
			controlShapeId: string;
		};
		expect(evidence.deletedShapeIds).toHaveLength(4);
		const saved = await editVsdx(
			bytes,
			evidence.deletedShapeIds.map((shapeId) => ({
				...remove,
				pageId: original.pages[0]!.id,
				shapeId,
			})),
		);
		const result = await parseVsdx(saved.bytes);
		expect(result.pages[0]!.shapes).toEqual(native.pages[0]!.shapes);
		expect(result.pages[0]!.shapes.map((shape) => shape.id)).toEqual([evidence.controlShapeId]);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
	},
);
