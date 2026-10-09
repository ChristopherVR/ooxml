import { expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture } from './test-fixtures';
import { attribute, children } from './sheet';
import { snapshotEdits } from './ui/edit-commands';

const box = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width: 1,
	height: 1,
});
const connector = (connect: { begin?: string; end?: string }, endX = 9, endY = 9): VisioEdit => ({
	type: 'create-line',
	pageId: '0',
	shapeId: '3',
	beginX: 0,
	beginY: 0,
	endX,
	endY,
	connect,
});
async function page(bytes: Uint8Array): Promise<Element> {
	const pkg = await VisioPackage.open(bytes);
	return pkg.readXml('visio/pages/page1.xml');
}
const shapeCells = (root: Element, id: string) => {
	const shape = Array.from(root.getElementsByTagName('Shape')).find(
		(node) => attribute(node, 'ID') === id,
	)!;
	return new Map(children(shape, 'Cell').map((node) => [attribute(node, 'N')!, node]));
};
const value = (root: Element, id: string, name: string) =>
	Number(attribute(shapeCells(root, id).get(name), 'V'));
const connects = (root: Element) =>
	children(children(root, 'Connects')[0], 'Connect').map((node) =>
		['FromSheet', 'FromCell', 'FromPart', 'ToSheet', 'ToCell', 'ToPart'].map((name) =>
			attribute(node, name),
		),
	);
async function glued(): Promise<Uint8Array> {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	const shapes = await editVsdx(blank, [box('1', 2, 2), box('2', 5, 2)]);
	return (await editVsdx(shapes.bytes, [connector({ begin: '1', end: '2' })])).bytes;
}

it('creates a native dynamically glued connector between side midpoints', async () => {
	const bytes = await glued();
	const root = await page(bytes);
	expect([value(root, '3', 'BeginX'), value(root, '3', 'BeginY')]).toEqual([2.5, 2]);
	expect([value(root, '3', 'EndX'), value(root, '3', 'EndY')]).toEqual([4.5, 2]);
	expect(value(root, '3', 'Width')).toBeCloseTo(2, 12);
	expect(attribute(shapeCells(root, '3').get('BeginX'), 'F')).toBe(
		'_WALKGLUE(BegTrigger,EndTrigger,WalkPreference)',
	);
	expect(attribute(shapeCells(root, '3').get('EndTrigger'), 'F')).toBe(
		'_XFTRIGGER(Sheet.2!EventXFMod)',
	);
	expect(value(root, '3', 'ObjType')).toBe(2);
	expect(connects(root)).toEqual([
		['3', 'BeginX', '9', '1', 'PinX', '3'],
		['3', 'EndX', '12', '2', 'PinX', '3'],
	]);
	const model = await parseVsdx(bytes);
	expect(model.pages[0]!.connectors).toEqual([
		expect.objectContaining({ fromShapeId: '3', toShapeId: '1', fromCell: 'BeginX' }),
		expect.objectContaining({ fromShapeId: '3', toShapeId: '2', fromCell: 'EndX' }),
	]);
	expect(model.pages[0]!.shapes[2]!.kind).toBe('connector');
});

it('reroutes glued ends when a shape moves or resizes', async () => {
	const moved = await editVsdx(await glued(), [
		{ type: 'move-shape', pageId: '0', shapeId: '2', x: 5, y: 4 },
	]);
	let root = await page(moved.bytes);
	expect([value(root, '3', 'BeginX'), value(root, '3', 'BeginY')]).toEqual([2.5, 2]);
	expect([value(root, '3', 'EndX'), value(root, '3', 'EndY')]).toEqual([4.5, 4]);
	expect(value(root, '3', 'PinX')).toBeCloseTo(3.5, 12);
	expect(value(root, '3', 'Width')).toBeCloseTo(Math.hypot(2, 2), 12);
	expect(connects(root)).toHaveLength(2);
	const resized = await editVsdx(moved.bytes, [
		{
			type: 'resize-shape',
			pageId: '0',
			shapeId: '2',
			width: 3,
			height: 1,
			anchor: { x: 1, y: 0.5 },
		},
	]);
	root = await page(resized.bytes);
	expect(value(root, '2', 'PinX')).toBe(4);
	// Both ends walk: the top midpoint of 1 and the left midpoint of the wider 2 are now closest.
	expect([value(root, '3', 'BeginX'), value(root, '3', 'BeginY')]).toEqual([2, 2.5]);
	expect([value(root, '3', 'EndX'), value(root, '3', 'EndY')]).toEqual([2.5, 4]);
	const plain = await editVsdx(resized.bytes, [
		{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 2, height: 1 },
	]);
	root = await page(plain.bytes);
	expect(value(root, '3', 'BeginX')).toBe(2);
	expect(value(root, '3', 'BeginY')).toBe(2.5);
	const parsed = await parseVsdx(plain.bytes);
	expect(parsed.pages[0]!.connectors).toHaveLength(2);
});

it('unglues a moved connector and heals glue when a glued shape is deleted', async () => {
	const bytes = await glued();
	const detached = await editVsdx(bytes, [
		{ type: 'move-line-endpoint', pageId: '0', shapeId: '3', endpoint: 'end', x: 6, y: 6 },
	]);
	let root = await page(detached.bytes);
	expect(connects(root)).toEqual([['3', 'BeginX', '9', '1', 'PinX', '3']]);
	expect(attribute(shapeCells(root, '3').get('EndX'), 'F')).toBeUndefined();
	const follows = await editVsdx(detached.bytes, [
		{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2, y: 5 },
	]);
	root = await page(follows.bytes);
	expect([value(root, '3', 'EndX'), value(root, '3', 'EndY')]).toEqual([6, 6]);
	expect(value(root, '3', 'BeginX')).toBe(2.5);
	expect(value(root, '3', 'BeginY')).toBe(5);

	const deleted = await editVsdx(bytes, [{ type: 'delete-shape', pageId: '0', shapeId: '2' }]);
	root = await page(deleted.bytes);
	expect(connects(root)).toEqual([['3', 'BeginX', '9', '1', 'PinX', '3']]);
	expect(attribute(shapeCells(root, '3').get('EndY'), 'F')).toBeUndefined();
	expect(value(root, '3', 'EndX')).toBe(4.5);
	const both = await editVsdx(deleted.bytes, [{ type: 'delete-shape', pageId: '0', shapeId: '3' }]);
	root = await page(both.bytes);
	expect(children(root, 'Connects')).toHaveLength(0);
	const mixed = await editVsdx(bytes, [
		{ type: 'move-shape', pageId: '0', shapeId: '1', x: 1, y: 1 },
		{ type: 'delete-shape', pageId: '0', shapeId: '1' },
	]);
	expect(connects(await page(mixed.bytes))).toEqual([['3', 'EndX', '12', '2', 'PinX', '3']]);
});

it('leaves an end on empty canvas unglued and validates glue targets', async () => {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	const one = (await editVsdx(blank, [box('1', 2, 2)])).bytes;
	const saved = await editVsdx(one, [connector({ begin: '1' }, 6, 2)]);
	const root = await page(saved.bytes);
	expect(connects(root)).toEqual([['3', 'BeginX', '9', '1', 'PinX', '3']]);
	expect([value(root, '3', 'BeginX'), value(root, '3', 'EndX')]).toEqual([2.5, 6]);
	expect(value(root, '3', 'EndTrigger')).toBe(0);
	for (const connect of [{ begin: '3' }, { begin: '1', end: '1' }, { begin: '01' }, { end: '7' }])
		await expect(editVsdx(one, [connector(connect)])).rejects.toThrow();
	await expect(
		editVsdx(saved.bytes, [{ ...connector({ begin: '3' }), shapeId: '4' } as VisioEdit]),
	).rejects.toThrow();
	const [copied] = snapshotEdits([connector({ begin: '1', extra: 'x' } as never)]);
	expect((copied as Extract<VisioEdit, { type: 'create-line' }>).connect).toEqual({ begin: '1' });
});

it('keeps glue through formatting, text and stacking edits, and unglues a dragged connector', async () => {
	const bytes = await glued();
	for (const edit of [
		{ type: 'format-shape', pageId: '0', shapeId: '3', lineColor: '#FF0000' },
		{ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#00FF00' },
		{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'Start' },
		{ type: 'reorder-shape', pageId: '0', shapeId: '3', order: 'back' },
		{ type: 'rotate-shape', pageId: '0', shapeId: '2', angle: Math.PI / 2 },
	] as VisioEdit[])
		expect(connects(await page((await editVsdx(bytes, [edit])).bytes))).toHaveLength(2);
	const dragged = await editVsdx(bytes, [
		{ type: 'move-shape', pageId: '0', shapeId: '3', x: 3.5, y: 5 },
	]);
	const root = await page(dragged.bytes);
	expect(children(root, 'Connects')).toHaveLength(0);
	expect([value(root, '3', 'BeginY'), value(root, '3', 'EndY')]).toEqual([5, 5]);
});
