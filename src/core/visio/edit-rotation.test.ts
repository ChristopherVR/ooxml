import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture, cell, shape, rectangle } from './test-fixtures';
import { attribute, children } from './sheet';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';
const rotate = (angle = Math.PI / 6): Extract<VisioEdit, { type: 'rotate-shape' }> => ({
	type: 'rotate-shape',
	pageId: '0',
	shapeId: '1',
	angle,
});
const target = (extra = '', attrs = '') =>
	shape(
		'1',
		cell('Width', 2) + cell('Height', 1) + cell('PinX', 2) + cell('PinY', 3) + rectangle + extra,
		attrs,
	);

it('rotates about the existing pin and recalculates angular dependencies', async () => {
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${target(cell('Angle', 0) + cell('TxtAngle', 0, 'Angle*0.5'))}</Shapes>`,
			},
		],
	});
	const result = await editVsdx(bytes, [rotate(Math.PI / 2)]);
	const model = await parseVsdx(result.bytes);
	const transform = model.pages[0]!.shapes[0]!.transform;
	expect(transform[0]).toBeCloseTo(0, 12);
	expect(transform[1]).toBeCloseTo(1, 12);
	expect(transform[2]).toBeCloseTo(-1, 12);
	expect(transform[3]).toBeCloseTo(0, 12);
	expect(transform[4]).toBeCloseTo(2.5, 12);
	expect(transform[5]).toBeCloseTo(2, 12);
	const root = await (
		await VisioPackage.open(result.bytes)
	).readXml('visio/pages/page1.xml', 'PageContents');
	const cells = children(children(children(root, 'Shapes')[0], 'Shape')[0], 'Cell');
	expect(
		Number(
			attribute(
				cells.find((node) => attribute(node, 'N') === 'TxtAngle'),
				'V',
			),
		),
	).toBeCloseTo(Math.PI / 4, 12);
	const before = await VisioPackage.open(bytes),
		after = await VisioPackage.open(result.bytes);
	for (const path of before.paths())
		if (path !== 'visio/pages/page1.xml')
			expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
});
it('keeps a zero-angle no-op byte preserving and snapshots independent angular commands', async () => {
	const bytes = await fixture({ pages: [{ id: '0', contents: `<Shapes>${target()}</Shapes>` }] });
	expect((await editVsdx(bytes, [rotate(0)])).bytes).toEqual(bytes);
	const command = rotate();
	const copy = snapshotEdits([{ ...command, extra: new Map() } as VisioEdit]);
	command.angle = 4;
	expect(copy).toEqual([rotate()]);
	const page = (await parseVsdx(bytes)).pages[0]!;
	expect(visioPageEditToDrawing({ ...page, drawingToPageScale: 3 }, rotate())).toEqual(rotate());
	for (const value of [Infinity, NaN, 1e7])
		await expect(editVsdx(bytes, [rotate(value)])).rejects.toThrow();
});
for (const [name, extra, document] of [
	['local rotation lock', cell('LockRotate', 1), ''],
	['guarded angle', cell('Angle', 0, 'GUARD(0rad)'), ''],
	['inconsistent dependent angle', cell('Angle', 0, 'Width/1in'), ''],
	['invalid angle unit', '<Cell N="Angle" V="0" U="IN"/>', ''],
	[
		'inherited rotation lock',
		'',
		`<StyleSheets><StyleSheet ID="0">${cell('LockRotate', 1)}</StyleSheet></StyleSheets>`,
	],
	['unknown affected dependency', cell('Angle', 0) + cell('TxtAngle', 0, 'NOW()+Angle'), ''],
] as const)
	it(`refuses ${name}`, async () => {
		const bytes = await fixture({
			document,
			pages: [{ id: '0', contents: `<Shapes>${target(extra)}</Shapes>` }],
		});
		await expect(editVsdx(bytes, [rotate()])).rejects.toThrow();
	});
it('replaces a proven formula on an explicit same-value angle assignment', async () => {
	const source = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${target(cell('Angle', Math.PI / 6, 'Width/1in*15deg') + cell('TxtAngle', Math.PI / 12, 'Angle*0.5'))}</Shapes>`,
			},
		],
	});
	const edited = await editVsdx(source, [rotate(Math.PI / 6)]);
	expect(edited.changedParts).toEqual(['visio/pages/page1.xml']);
	expect((await parseVsdx(edited.bytes)).pages[0]!.shapes[0]!.transform).toEqual(
		(await parseVsdx(source)).pages[0]!.shapes[0]!.transform,
	);
	const root = await (
		await VisioPackage.open(edited.bytes)
	).readXml('visio/pages/page1.xml', 'PageContents');
	const nodes = children(children(children(root, 'Shapes')[0], 'Shape')[0], 'Cell');
	expect(
		attribute(
			nodes.find((node) => attribute(node, 'N') === 'Angle'),
			'F',
		),
	).toBeUndefined();
	expect(
		attribute(
			nodes.find((node) => attribute(node, 'N') === 'TxtAngle'),
			'F',
		),
	).toBe('Angle*0.5');
	expect((await editVsdx(edited.bytes, [rotate(Math.PI / 6)])).bytes).toEqual(edited.bytes);
});
it('converts a cached degree angle to native radians on rotation', async () => {
	const bytes = await fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${target('<Cell N="Angle" V="45" U="DEG"/>')}</Shapes>` },
		],
	});
	const edited = await editVsdx(bytes, [rotate(-Math.PI / 4)]);
	expect((await parseVsdx(edited.bytes)).pages[0]!.shapes[0]!.transform[1]).toBeCloseTo(
		-Math.SQRT1_2,
		12,
	);
});
for (const variable of [
	'VISIO_NATIVE_ROTATE_DIR',
	'VISIO_NATIVE_ROTATE_DOUBLE_DIR',
	'VISIO_NATIVE_ROTATE_HALF_DIR',
	'VISIO_NATIVE_ROTATE_TRIPLE_DIR',
	'VISIO_NATIVE_ROTATE_DEPENDENT_DIR',
	'VISIO_NATIVE_ROTATE_DEPENDENT_SAME_DIR',
]) {
	const directory = process.env[variable];
	for (const kind of ['rectangle', 'ellipse'])
		it.skipIf(!directory)(`matches native ${kind} rotation (${variable})`, async () => {
			const source = await readFile(
				join(
					directory!,
					variable.includes('_DEPENDENT_') ? 'rotation-source.vsdx' : 'ellipse-edited.vsdx',
				),
			);
			const native = await parseVsdx(await readFile(join(directory!, 'rotated.vsdx')));
			const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
				rotated: Record<
					string,
					{ shapeId: string; cells: Record<string, { value: number }>; transform: number[] }
				>;
			};
			const reference = evidence.rotated[kind]!;
			const pageId = native.pages[0]!.id;
			const edited = await editVsdx(source, [
				{
					type: 'rotate-shape',
					pageId,
					shapeId: reference.shapeId,
					angle: reference.cells.Angle!.value,
				},
			]);
			if (variable.includes('_DEPENDENT_')) {
				const before = await VisioPackage.open(source),
					after = await VisioPackage.open(edited.bytes);
				const getAngle = async (pkg: VisioPackage) => {
					const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
					const node = children(children(root, 'Shapes')[0], 'Shape').find(
						(node) => attribute(node, 'ID') === reference.shapeId,
					)!;
					return children(node, 'Cell').find((node) => attribute(node, 'N') === 'Angle')!;
				};
				expect(attribute(await getAngle(before), 'F')).toContain('Width');
				expect(attribute(await getAngle(after), 'F')).toBeUndefined();
				if (variable.includes('_SAME_'))
					expect(attribute(await getAngle(after), 'V')).toBe(
						attribute(await getAngle(before), 'V'),
					);
				expect(edited.changedParts).toContain('visio/pages/page1.xml');
				for (const path of before.paths())
					if (path !== 'visio/pages/page1.xml')
						expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
			}
			const actual = (await parseVsdx(edited.bytes)).pages[0]!.shapes.find(
				(shape) => shape.id === reference.shapeId,
			)!;
			const expected = native.pages[0]!.shapes.find((shape) => shape.id === reference.shapeId)!;
			expect(actual.geometry).toEqual(expected.geometry);
			expect(actual.style).toEqual(expected.style);
			const ratio = native.pages[0]!.drawingToPageScale ?? 1;
			for (let i = 0; i < 6; i++) {
				expect(actual.transform[i]).toBeCloseTo(expected.transform[i]!, 12);
				expect(actual.transform[i]).toBeCloseTo(reference.transform[i]! * (i >= 4 ? ratio : 1), 12);
			}
		});
}

it('refuses connector, group and glued rotation outside the proven local shape scope', async () => {
	const line = shape(
		'1',
		cell('OneD', 1) + cell('Width', 2) + cell('Height', 0) + cell('PinX', 2) + cell('PinY', 3),
	);
	const group = target(`<Shapes>${shape('2')}</Shapes>`);
	const glued = `${target()}${shape('2')}`;
	for (const contents of [
		`<Shapes>${line}</Shapes>`,
		`<Shapes>${group}</Shapes>`,
		`<Shapes>${glued}</Shapes><Connects><Connect FromSheet="2" ToSheet="1" FromCell="BeginX" ToCell="PinX"/></Connects>`,
	])
		await expect(
			editVsdx(await fixture({ pages: [{ id: '0', contents }] }), [rotate()]),
		).rejects.toThrow();
});
