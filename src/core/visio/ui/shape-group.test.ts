import { it, expect } from 'vitest';
import { parseVsdx } from '../parser';
import { editVsdx } from '../edit';
import { fixture } from '../test-fixtures';
import { createVsdx } from '../create-document';
import { visioGroupCommand, visioUngroupCommand, visioMovableGroup } from './shape-group';
import { visioMovementShape } from './shape-move';
import { visioResizeShape } from './shape-resize';

it('plans Group for two or more admissible top-level shapes in page order', async () => {
	const blank = await createVsdx();
	const pageId = (await parseVsdx(blank)).pages[0]!.id;
	const drawn = await editVsdx(blank, [
		{ type: 'create-rectangle', pageId, shapeId: '1', x: 1, y: 1, width: 1, height: 1 },
		{ type: 'create-ellipse', pageId, shapeId: '2', x: 3, y: 1, width: 1, height: 1 },
		{ type: 'create-line', pageId, shapeId: '3', beginX: 0, beginY: 0, endX: 1, endY: 1 },
	]);
	const page = (await parseVsdx(drawn.bytes)).pages[0]!;
	expect(visioGroupCommand(page, ['2', '1'])).toEqual({
		type: 'group-shapes',
		pageId,
		shapeId: '4',
		memberIds: ['1', '2'],
	});
	for (const ids of [['1'], ['1', '1'], ['1', '3'], ['1', '9']])
		expect(visioGroupCommand(page, ids)).toBeUndefined();
	expect(visioUngroupCommand(page, ['1'])).toBeUndefined();
	const grouped = await editVsdx(drawn.bytes, [visioGroupCommand(page, ['1', '2'])!]);
	const after = (await parseVsdx(grouped.bytes)).pages[0]!;
	expect(visioUngroupCommand(after, ['4'])).toEqual({
		type: 'ungroup-shape',
		pageId,
		shapeId: '4',
	});
	expect(visioUngroupCommand(after, ['4', '3'])).toBeUndefined();
	expect(visioMovableGroup(after, '4')).toBeDefined();
	expect(visioMovementShape(after, '4')).toBeDefined();
	expect(visioResizeShape(after, '4')).toBeUndefined();
});

it('keeps glued or master-based groups out of the pointer move scope', async () => {
	const group = `<Shape ID="1" Type="Group"><Cell N="PinX" V="2"/><Cell N="PinY" V="2"/><Cell N="Width" V="2"/><Cell N="Height" V="2"/><Shapes><Shape ID="2" Type="Shape" Master="7"><Cell N="Width" V="1"/><Cell N="Height" V="1"/></Shape></Shapes></Shape>`;
	const page = (
		await parseVsdx(await fixture({ pages: [{ id: '0', contents: `<Shapes>${group}</Shapes>` }] }))
	).pages[0]!;
	expect(visioMovableGroup(page, '1')).toBeUndefined();
	expect(visioUngroupCommand(page, ['1'])).toBeUndefined();
});
