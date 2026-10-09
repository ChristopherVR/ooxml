import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture } from './test-fixtures';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';
import { attribute, children } from './sheet';
import { buildXml } from '../xml';
import type { VisioPathCreateEdit } from './edit-path-commands';

const path = (
	segments: VisioPathCreateEdit['segments'],
	closed?: boolean,
): VisioPathCreateEdit => ({
	type: 'create-path',
	pageId: '0',
	shapeId: '2',
	x: 1,
	y: 1,
	segments: structuredClone(segments),
	...(closed === undefined ? {} : { closed }),
});
const quarter: VisioPathCreateEdit['segments'] = [
	{ kind: 'arc', x: 5, y: 3, a: 1 + 4 * Math.SQRT1_2, b: 3 - 2 * Math.SQRT1_2, ratio: 2 },
];
const curve: VisioPathCreateEdit['segments'] = [
	{ kind: 'cubic', x1: 1, y1: 3, x2: 3, y2: 3, x: 3, y: 1 },
	{ kind: 'line', x: 4, y: 2 },
];
const rowsOf = async (bytes: Uint8Array, shapeId = '2') => {
	const pkg = await VisioPackage.open(bytes),
		root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
	const shape = children(children(root, 'Shapes')[0], 'Shape').find(
		(node) => attribute(node, 'ID') === shapeId,
	)!;
	return { root, shape, rows: children(children(shape, 'Section')[0], 'Row') };
};

it('creates a relative quarter-ellipse arc that renders, moves, resizes and rotates', async () => {
	const created = await editVsdx(await fixture(), [path(quarter)]);
	const { rows, shape } = await rowsOf(created.bytes);
	expect(rows.map((row) => attribute(row, 'T'))).toEqual(['RelMoveTo', 'RelEllipticalArcTo']);
	expect(
		attribute(
			children(rows[1]!, 'Cell').find((cell) => attribute(cell, 'N') === 'D'),
			'F',
		),
	).toBe('Width/Height*1');
	expect(
		attribute(
			children(shape, 'Cell').find((c) => attribute(c, 'N') === 'Width'),
			'V',
		),
	).toBe('4');
	let model = await parseVsdx(created.bytes);
	let arc = model.pages[0]!.shapes.find((item) => item.id === '2')!;
	expect(arc.width).toBe(4);
	expect(arc.height).toBe(2);
	expect(arc.geometry[0]!.path).toBe('M 0 0 A 4 2 0 0 1 4 2');
	expect(arc.geometry[0]!.fill).toBe(false);
	const edited = await editVsdx(created.bytes, [
		{ type: 'resize-shape', pageId: '0', shapeId: '2', width: 2, height: 4 },
		{ type: 'move-shape', pageId: '0', shapeId: '2', x: 6, y: 6 },
		{ type: 'rotate-shape', pageId: '0', shapeId: '2', angle: Math.PI / 2 },
	]);
	model = await parseVsdx(edited.bytes);
	arc = model.pages[0]!.shapes.find((item) => item.id === '2')!;
	expect(arc.width).toBe(2);
	expect(arc.height).toBe(4);
	// The recalculated ratio follows the new frame: still a quarter of the frame's ellipse.
	expect(arc.geometry[0]!.path).toBe('M 0 0 A 2 4 0 0 1 2 4');
	const { rows: resized } = await rowsOf(edited.bytes);
	expect(
		Number(
			attribute(
				children(resized[1]!, 'Cell').find((c) => attribute(c, 'N') === 'D'),
				'V',
			),
		),
	).toBeCloseTo(0.5, 12);
	const removed = await editVsdx(edited.bytes, [
		{ type: 'delete-shape', pageId: '0', shapeId: '2' },
	]);
	expect((await parseVsdx(removed.bytes)).pages[0]!.shapes.map((item) => item.id)).toEqual(['1']);
});

it('creates freeform cubic paths, open unfilled and closed filled, and resizes them', async () => {
	const created = await editVsdx(await fixture(), [path(curve)]);
	const { rows } = await rowsOf(created.bytes);
	expect(rows.map((row) => attribute(row, 'T'))).toEqual(['RelMoveTo', 'RelCubBezTo', 'RelLineTo']);
	const shape = (await parseVsdx(created.bytes)).pages[0]!.shapes.find((item) => item.id === '2')!;
	expect(shape.geometry[0]!.fill).toBe(false);
	expect(shape.geometry[0]!.path).toMatch(/^M 0 \d.* C .* L /);
	const resized = await editVsdx(created.bytes, [
		{
			type: 'resize-shape',
			pageId: '0',
			shapeId: '2',
			width: shape.width * 2,
			height: shape.height,
		},
	]);
	const wide = (await parseVsdx(resized.bytes)).pages[0]!.shapes.find((item) => item.id === '2')!;
	expect(wide.width).toBeCloseTo(shape.width * 2, 12);
	const closed = await editVsdx(await fixture(), [
		path(
			[
				{ kind: 'line', x: 3, y: 1 },
				{ kind: 'cubic', x1: 3, y1: 2, x2: 2, y2: 3, x: 1, y: 3 },
				{ kind: 'line', x: 1, y: 1 },
			],
			true,
		),
	]);
	const filled = (await parseVsdx(closed.bytes)).pages[0]!.shapes.find((item) => item.id === '2')!;
	expect(filled.geometry[0]!.fill).toBe(true);
	expect(filled.geometry[0]!.path).toBe('M 0 0 L 2 0 C 2 1 1 2 0 2 L 0 0');
});

it('snapshots, scales and validates path creation', async () => {
	const command = path(quarter);
	const copied = snapshotEdits([{ ...command, extra: 1 } as VisioEdit]);
	(command.segments[0] as { x: number }).x = 99;
	expect(copied).toEqual([path(quarter)]);
	const page = (await parseVsdx(await fixture())).pages[0]!;
	const scaled = visioPageEditToDrawing({ ...page, drawingToPageScale: 2 }, path(curve));
	expect(scaled).toMatchObject({
		x: 0.5,
		y: 0.5,
		segments: [{ x1: 0.5, y1: 1.5, x: 1.5 }, { x: 2 }],
	});
	for (const edit of [
		path([]),
		path([{ kind: 'line', x: 4, y: 1 }]),
		path([{ kind: 'line', x: 4, y: 2 }], true),
		path([{ kind: 'arc', x: 4, y: 2, a: 2, b: 2, ratio: 0 }]),
		path([{ kind: 'bogus', x: 1, y: 1 } as never]),
		{ ...path(curve), shapeId: '1' },
	])
		await expect(editVsdx(await fixture(), [edit as VisioEdit])).rejects.toThrow();
});

it('refuses to resize a relative arc whose ratio lacks the Width/Height proof', async () => {
	const created = await editVsdx(await fixture(), [path(quarter)]);
	const zip = await JSZip.loadAsync(created.bytes);
	const { root, rows } = await rowsOf(created.bytes);
	const ratio = children(rows[1]!, 'Cell').find((cell) => attribute(cell, 'N') === 'D')!;
	ratio.removeAttribute('F');
	zip.file('visio/pages/page1.xml', buildXml(root));
	const modified = await zip.generateAsync({ type: 'uint8array' });
	await expect(
		editVsdx(modified, [{ type: 'resize-shape', pageId: '0', shapeId: '2', width: 2, height: 2 }]),
	).rejects.toMatchObject({ code: 'UNSUPPORTED_GEOMETRY_EDIT' });
	// Moving needs no geometry proof.
	await expect(
		editVsdx(modified, [{ type: 'move-shape', pageId: '0', shapeId: '2', x: 4, y: 4 }]),
	).resolves.toBeDefined();
});
