import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture } from './test-fixtures';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';
import { attribute, children } from './sheet';
import { buildXml } from '../xml';
import JSZip from 'jszip';
const create = (): Extract<VisioEdit, { type: 'create-ellipse' }> => ({
	type: 'create-ellipse',
	pageId: '0',
	shapeId: '2',
	x: 2,
	y: 3,
	width: 2,
	height: 1,
});

it('creates, moves, resizes and deletes an ellipse through shared editing and recalc', async () => {
	const source = await fixture();
	const created = await editVsdx(source, [create()]);
	const resized = await editVsdx(created.bytes, [
		{ type: 'resize-shape', pageId: '0', shapeId: '2', width: 4, height: 2 },
		{ type: 'move-shape', pageId: '0', shapeId: '2', x: 5, y: 6 },
	]);
	const model = await parseVsdx(resized.bytes),
		ellipse = model.pages[0]!.shapes.find((shape) => shape.id === '2')!;
	expect(ellipse.geometry[0]!.path).toBe('M 4 1 A 2 1 0 0 1 0 1 A 2 1 0 0 1 4 1 Z');
	expect(ellipse.transform.map((value) => value || 0)).toEqual([1, 0, 0, 1, 3, 5]);
	expect(ellipse.width).toBe(4);
	expect(ellipse.height).toBe(2);
	const before = await VisioPackage.open(source),
		after = await VisioPackage.open(resized.bytes);
	for (const path of before.paths())
		if (path !== 'visio/pages/page1.xml')
			expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
	const removed = await editVsdx(resized.bytes, [
		{ type: 'delete-shape', pageId: '0', shapeId: '2' },
	]);
	expect((await parseVsdx(removed.bytes)).pages[0]!.shapes.map((shape) => shape.id)).toEqual(['1']);
});

it('snapshots and scales ellipse creation through the shared box command path', async () => {
	const command = create();
	const copied = snapshotEdits([{ ...command, extra: new Map() } as VisioEdit]);
	command.width = 99;
	expect(copied).toEqual([create()]);
	const page = (await parseVsdx(await fixture())).pages[0]!;
	expect(visioPageEditToDrawing({ ...page, drawingToPageScale: 2 }, create())).toEqual({
		...create(),
		x: 1,
		y: 1.5,
		width: 1,
		height: 0.5,
	});
	for (const edit of [
		{ ...create(), width: 0 },
		{ ...create(), height: Infinity },
		{ ...create(), shapeId: '1' },
	])
		await expect(editVsdx(await fixture(), [edit])).rejects.toThrow();
});

it('refuses changed ellipse axes that lack the native resize proof', async () => {
	const saved = await editVsdx(await fixture(), [create()]);
	const zip = await JSZip.loadAsync(saved.bytes);
	const pkg = await VisioPackage.open(saved.bytes),
		root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
	const shape = children(children(root, 'Shapes')[0], 'Shape').find(
		(shape) => attribute(shape, 'ID') === '2',
	)!;
	const row = children(children(shape, 'Section')[0], 'Row')[0]!;
	const a = children(row, 'Cell').find((cell) => attribute(cell, 'N') === 'A')!;
	a.setAttribute('F', 'Width*2');
	a.setAttribute('V', '4');
	zip.file('visio/pages/page1.xml', buildXml(root));
	const modified = await zip.generateAsync({ type: 'uint8array' });
	await expect(
		editVsdx(modified, [{ type: 'resize-shape', pageId: '0', shapeId: '2', width: 4, height: 2 }]),
	).rejects.toMatchObject({ code: 'UNSUPPORTED_GEOMETRY_EDIT' });
});

for (const variable of [
	'VISIO_NATIVE_DRAW_DEFAULTS_DIR',
	'VISIO_NATIVE_DRAW_CUSTOM_DEFAULTS_DIR',
]) {
	const directory = process.env[variable];
	it.skipIf(!directory)(`matches native ellipse creation and editing (${variable})`, async () => {
		const bytes = await readFile(join(directory!, 'original.vsdx'));
		const original = await parseVsdx(bytes);
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			ellipse: {
				shapeId: string;
				cells: Record<string, { value: number }>;
				transform: number[];
				editedCells: Record<string, { value: number }>;
				editedTransform: number[];
			};
		};
		const target = evidence.ellipse;
		const cleared = await editVsdx(bytes, [
			{ type: 'delete-shape', pageId: original.pages[0]!.id, shapeId: target.shapeId },
		]);
		const made = await editVsdx(cleared.bytes, [
			{
				type: 'create-ellipse',
				pageId: original.pages[0]!.id,
				shapeId: target.shapeId,
				x: target.cells.PinX!.value,
				y: target.cells.PinY!.value,
				width: target.cells.Width!.value,
				height: target.cells.Height!.value,
			},
		]);
		const edited = await editVsdx(made.bytes, [
			{
				type: 'resize-shape',
				pageId: original.pages[0]!.id,
				shapeId: target.shapeId,
				width: target.editedCells.Width!.value,
				height: target.editedCells.Height!.value,
			},
			{
				type: 'move-shape',
				pageId: original.pages[0]!.id,
				shapeId: target.shapeId,
				x: target.editedCells.PinX!.value,
				y: target.editedCells.PinY!.value,
			},
		]);
		const nativeEdited = await parseVsdx(await readFile(join(directory!, 'ellipse-edited.vsdx')));
		for (const [result, native, pose] of [
			[made, original, target.transform],
			[edited, nativeEdited, target.editedTransform],
		] as const) {
			const model = await parseVsdx(result.bytes),
				actual = model.pages[0]!.shapes.find((shape) => shape.id === target.shapeId)!;
			const expected = native.pages[0]!.shapes.find((shape) => shape.id === target.shapeId)!;
			expect(actual.geometry).toEqual(expected.geometry);
			expect(actual.style).toEqual(expected.style);
			for (let i = 0; i < 6; i++) expect(actual.transform[i]).toBeCloseTo(pose[i]!, 12);
		}
	});
}
