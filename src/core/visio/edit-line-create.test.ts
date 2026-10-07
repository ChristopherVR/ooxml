import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture, shape, cell, rectangle } from './test-fixtures';
import { attribute, children } from './sheet';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';
import { visioNextShapeId } from './ui/shape-id';

const create = (endX = 3, endY = 1.5): Extract<VisioEdit, { type: 'create-line' }> => ({
	type: 'create-line',
	pageId: '0',
	shapeId: '2',
	beginX: 1,
	beginY: 1.5,
	endX,
	endY,
});
const source = () =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 1) + cell('Height', 1) + rectangle)}</Shapes>`,
			},
		],
	});

it.each([
	[3, 1.5],
	[2.732050807568877, 2.5],
	[1, 3.5],
	[-1, 1.5],
])('creates editable native straight geometry (%s, %s)', async (x, y) => {
	const bytes = await source();
	const saved = await editVsdx(bytes, [create(x, y)]);
	const model = await parseVsdx(saved.bytes);
	const line = model.pages[0]!.shapes[1]!;
	expect(line.kind).toBe('connector');
	expect(line.height).toBe(0);
	expect(line.width).toBeCloseTo(Math.hypot(x - 1, y - 1.5), 12);
	expect(line.transform[4]).toBeCloseTo(1, 12);
	expect(line.transform[5]).toBeCloseTo(1.5, 12);
	expect(line.geometry[0]).toMatchObject({ fill: false, stroke: true });
	const moved = await editVsdx(saved.bytes, [
		{ type: 'move-line-endpoint', pageId: '0', shapeId: '2', endpoint: 'end', x: 4, y: 3 },
	]);
	expect((await parseVsdx(moved.bytes)).pages[0]!.shapes[1]!.width).toBeCloseTo(
		Math.hypot(3, 1.5),
		12,
	);
	const deleted = await editVsdx(moved.bytes, [
		{ type: 'delete-shape', pageId: '0', shapeId: '2' },
	]);
	expect((await parseVsdx(deleted.bytes)).pages[0]!.shapes).toEqual(
		(await parseVsdx(bytes)).pages[0]!.shapes,
	);
	const before = await VisioPackage.open(bytes),
		after = await VisioPackage.open(saved.bytes);
	expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
	for (const path of before.paths())
		if (!saved.changedParts.includes(path))
			expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
});

it('bounds insertion, rejects coincident endpoints and shares collision admission', async () => {
	for (const edit of [
		create(1, 1.5),
		{ ...create(), endX: Infinity },
		{ ...create(), shapeId: '1' },
	])
		await expect(editVsdx(await source(), [edit])).rejects.toThrow();
});

it('snapshots line endpoints and converts physical page inches through shared scaling', async () => {
	const command = create();
	const copied = snapshotEdits([{ ...command, extra: new Map() } as VisioEdit]);
	command.endX = 99;
	expect(copied).toEqual([create()]);
	const page = (await parseVsdx(await source())).pages[0]!;
	expect(visioPageEditToDrawing({ ...page, drawingToPageScale: 2 }, create())).toEqual({
		...create(),
		beginX: 0.5,
		beginY: 0.75,
		endX: 1.5,
		endY: 0.75,
	});
});

it('allocates shape IDs across descendants through the native unsigned limit', async () => {
	const page = (await parseVsdx(await source())).pages[0]!;
	const child = { ...page.shapes[0]!, id: '4294967294' };
	page.shapes[0]!.children = [child];
	expect(visioNextShapeId(page)).toBe('4294967295');
	child.id = '4294967295';
	expect(() => visioNextShapeId(page)).toThrow('No shape IDs remain');
});
for (const variable of [
	'VISIO_NATIVE_LINE_CREATION_DIR',
	'VISIO_NATIVE_LINE_CREATION_HALF_DIR',
	'VISIO_NATIVE_LINE_CREATION_DOUBLE_DIR',
	'VISIO_NATIVE_LINE_CREATION_TRIPLE_DIR',
]) {
	const directory = process.env[variable];
	it.skipIf(!directory)(
		`matches native DrawLine geometry, defaults and cached transforms (${variable})`,
		async () => {
			const bytes = await readFile(join(directory!, 'original.vsdx'));
			const native = await parseVsdx(bytes);
			const before = await VisioPackage.open(bytes);
			const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
				pageScale: number;
				drawingScale: number;
				cases: {
					shapeId: string;
					before: Record<string, { value: number }>;
					beforeTransform: number[];
				}[];
			};
			const cleared = await editVsdx(
				bytes,
				evidence.cases.map((item) => ({
					type: 'delete-shape',
					pageId: native.pages[0]!.id,
					shapeId: item.shapeId,
				})),
			);
			const created = await editVsdx(
				cleared.bytes,
				evidence.cases.map((item) => ({
					type: 'create-line',
					pageId: native.pages[0]!.id,
					shapeId: item.shapeId,
					beginX: item.before.BeginX!.value,
					beginY: item.before.BeginY!.value,
					endX: item.before.EndX!.value,
					endY: item.before.EndY!.value,
				})),
			);
			const result = await parseVsdx(created.bytes);
			const pkg = await VisioPackage.open(created.bytes);
			expect(result.pages[0]!.drawingToPageScale).toBe(native.pages[0]!.drawingToPageScale);
			for (const path of before.paths())
				if (path !== 'visio/pages/page1.xml')
					expect(await pkg.readBytes(path)).toEqual(await before.readBytes(path));
			const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
			const ratio = evidence.pageScale / evidence.drawingScale;
			for (const item of evidence.cases) {
				const actual = result.pages[0]!.shapes.find((shape) => shape.id === item.shapeId)!;
				const expected = native.pages[0]!.shapes.find((shape) => shape.id === item.shapeId)!;
				for (let i = 0; i < 6; i++)
					expect(actual.transform[i]).toBeCloseTo(expected.transform[i]!, 12);
				expect(actual.width).toBeCloseTo(expected.width, 12);
				for (let i = 0; i < 6; i++)
					expect(actual.transform[i]).toBeCloseTo(
						item.beforeTransform[i]! * (i >= 4 ? ratio : 1),
						12,
					);
				expect(actual.width).toBeCloseTo(item.before.Width!.value * ratio, 12);
				expect(actual.geometry).toEqual(expected.geometry);
				expect(actual.style).toEqual(expected.style);
				const node = children(children(root, 'Shapes')[0], 'Shape').find(
					(node) => attribute(node, 'ID') === item.shapeId,
				)!;
				const cells = new Map(children(node, 'Cell').map((node) => [attribute(node, 'N'), node]));
				for (const [name, value] of Object.entries(item.before))
					expect(Number(attribute(cells.get(name), 'V')), name).toBeCloseTo(value.value, 12);
			}
		},
	);
}
