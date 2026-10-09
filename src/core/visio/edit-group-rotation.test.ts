import { it, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { composeAffine, IDENTITY_AFFINE, type AffineMatrix } from '../geometry/affine';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { children } from './sheet';
import { buildXml } from '../xml/xml';
import { fixture, cell, shape, rectangle } from './test-fixtures';
import type { VisioShape } from './model';
interface Tree {
	id: string;
	transform: number[];
	cells: Record<string, { value: number }>;
	children: Tree[];
}
function compare(
	shape: VisioShape,
	reference: Tree,
	ratio: number,
	parent: AffineMatrix = IDENTITY_AFFINE,
) {
	const world = composeAffine(parent, shape.transform);
	for (let i = 0; i < 6; i++)
		expect(world[i]).toBeCloseTo(reference.transform[i]! * (i >= 4 ? ratio : 1), 12);
	for (const child of shape.children)
		compare(
			child,
			reference.children.find((node) => node.id === child.id)!,
			ratio,
			world,
		);
}
for (const variable of ['VISIO_NATIVE_GROUP_ROTATE_DIR', 'VISIO_NATIVE_GROUP_ROTATE_NESTED_DIR'])
	it.skipIf(!process.env[variable])(
		`matches native parent and descendant matrices (${variable})`,
		async () => {
			const directory = process.env[variable]!;
			const evidence = JSON.parse(await readFile(join(directory, 'evidence.json'), 'utf8')) as {
				source: Tree;
				rotated: Tree;
			};
			const source = await readFile(join(directory, 'source.vsdx'));
			const before = await parseVsdx(source);
			const ratio = before.pages[0]!.drawingToPageScale ?? 1;
			compare(before.pages[0]!.shapes[0]!, evidence.source, ratio);
			const saved = await editVsdx(source, [
				{
					type: 'rotate-shape',
					pageId: '0',
					shapeId: evidence.source.id,
					angle: evidence.rotated.cells.Angle!.value,
				},
			]);
			const actual = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
			compare(actual, evidence.rotated, ratio);
			expect(actual.children).toEqual(before.pages[0]!.shapes[0]!.children);
			const original = await VisioPackage.open(source),
				edited = await VisioPackage.open(saved.bytes);
			const subtree = async (pkg: VisioPackage) => {
				const page = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
				return buildXml(children(children(children(page, 'Shapes')[0], 'Shape')[0], 'Shapes')[0]!);
			};
			expect(await subtree(edited)).toBe(await subtree(original));
			for (const path of original.paths())
				if (path !== 'visio/pages/page1.xml')
					expect(await edited.readBytes(path)).toEqual(await original.readBytes(path));
		},
	);
const transform = (angle = '') =>
	cell('Width', 2) +
	cell('Height', 1) +
	cell('PinX', 2) +
	cell('PinY', 3) +
	cell('LocPinX', 1) +
	cell('LocPinY', 0.5) +
	cell('Angle', 0, angle);
const group = (child: string, extra = '') =>
	shape('1', transform() + extra + `<Shapes>${child}</Shapes>`, 'Type="Group"');
const leaf = (angle = '', attrs = '') =>
	shape('2', transform(angle) + rectangle, 'Type="Shape" ' + attrs);
it('rotates a parent while preserving child local formulas and geometry', async () => {
	const source = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${group(leaf())}</Shapes>` }],
	});
	const before = (await parseVsdx(source)).pages[0]!.shapes[0]!;
	const saved = await editVsdx(source, [
		{ type: 'rotate-shape', pageId: '0', shapeId: '1', angle: Math.PI / 6 },
	]);
	const after = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
	expect(after.transform[0]).toBeCloseTo(Math.cos(Math.PI / 6), 12);
	expect(after.transform[1]).toBeCloseTo(Math.sin(Math.PI / 6), 12);
	expect(after.children).toEqual(before.children);
	expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
});
it.each([
	group(leaf('Sheet.1!Angle')),
	group(leaf('Inh')),
	group(
		shape('2', transform().replace('N="Angle"', 'N="Angle" E="1"') + rectangle, 'Type="Shape"'),
	),
	group(leaf('', 'Master="9"')),
	group(shape('2', transform() + '<ForeignData/>', 'Type="Shape"')),
	group(shape('2', transform() + cell('OneD', 1), 'Type="Shape"')),
	group(leaf(), cell('LockRotate', 1)),
])(
	'refuses unproven group dependencies/protection without source mutation (case %#)',
	async (contents) => {
		const source = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${contents}</Shapes>` }],
		});
		const original = source.slice();
		await expect(
			editVsdx(source, [{ type: 'rotate-shape', pageId: '0', shapeId: '1', angle: Math.PI / 6 }]),
		).rejects.toThrow();
		expect(source).toEqual(original);
	},
);
it('keeps group resizing and flips outside this rotation proof', async () => {
	const source = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${group(leaf())}</Shapes>` }],
	});
	for (const command of [
		{ type: 'resize-shape', width: 3, height: 2 },
		{ type: 'flip-shape', axis: 'horizontal' },
	] as const)
		await expect(editVsdx(source, [{ ...command, pageId: '0', shapeId: '1' }])).rejects.toThrow();
});

it('refuses glue on a descendant through the page Connects container', async () => {
	const source = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${group(leaf())}</Shapes><Connects><Connect FromSheet="2" ToSheet="1"/></Connects>`,
			},
		],
	});
	await expect(
		editVsdx(source, [{ type: 'rotate-shape', pageId: '0', shapeId: '1', angle: Math.PI / 6 }]),
	).rejects.toThrow();
});
