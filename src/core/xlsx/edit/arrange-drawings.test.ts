import { expect, it } from 'vitest';
import {
	anchorToPixelBox,
	createEditSession,
	createWorkbook,
	loadXlsx,
	pictureAnchorAt,
	saveXlsx,
} from '../index';
import type { DrawingObject } from '../model';

const shapeOf = (drawing: DrawingObject | undefined) =>
	drawing?.kind === 'unsupported' ? drawing.shape : undefined;
const reopen = async (workbook: ReturnType<typeof createWorkbook>) =>
	loadXlsx(await saveXlsx(workbook));

it('round-trips native preset shapes and multiline text boxes and edits their text', async () => {
	const book = createWorkbook();
	const session = createEditSession(book);
	const position = pictureAnchorAt({ row: 1, col: 1 }, 120, 60);
	session.addShape(0, position, { kind: 'preset', preset: 'hexagon', name: 'Hexagon' });
	session.addShape(0, position, { kind: 'textBox', name: 'Text Box', text: 'First < & >\nSecond' });
	expect(session.undo()).toBe(true);
	expect(book.sheets[0]?.drawings).toHaveLength(1);
	expect(session.redo()).toBe(true);
	const loaded = await reopen(book);
	expect(shapeOf(loaded.sheets[0]?.drawings[0])).toMatchObject({ geometry: 'hexagon' });
	expect(shapeOf(loaded.sheets[0]?.drawings[1])).toMatchObject({
		textBox: true,
		text: { text: 'First < & >\nSecond' },
	});
	const edit = createEditSession(loaded);
	edit.setTextBoxText(0, 1, 'New text');
	const edited = await reopen(loaded);
	expect(shapeOf(edited.sheets[0]?.drawings[1])?.text?.text).toBe('New text');
	expect(edit.undo()).toBe(true);
	expect(shapeOf(loaded.sheets[0]?.drawings[1])?.text?.text).toBe('First < & >\nSecond');
});

it('round-trips stacking, alignment, native groups, rotation and ungrouping', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	edit.addShape(0, pictureAnchorAt({ row: 1, col: 1 }, 110, 60), {
		kind: 'preset',
		preset: 'rect',
		name: 'First',
	});
	edit.addShape(0, pictureAnchorAt({ row: 4, col: 2 }, 90, 40), {
		kind: 'preset',
		preset: 'ellipse',
		name: 'Second',
	});
	edit.reorderDrawing(0, 0, 1);
	edit.alignDrawings(0, [0, 1], 'left');
	edit.rotateDrawing(0, 0, 90);
	const reopened = await reopen(book);
	const sheet = reopened.sheets[0]!;
	expect(sheet.drawings.map((item) => shapeOf(item)?.name)).toEqual(['Second', 'First']);
	expect(shapeOf(sheet.drawings[0])?.rotation).toBe(90);
	const x = sheet.drawings.map((item) => anchorToPixelBox(sheet, item.anchor).x);
	expect(x[0]).toBe(x[1]);
	const groupEdit = createEditSession(reopened);
	expect(groupEdit.groupDrawings(0, [0, 1], 'Group')).toBe(0);
	expect(groupEdit.undo()).toBe(true);
	expect(sheet.drawings).toHaveLength(2);
	expect(groupEdit.redo()).toBe(true);
	const grouped = await reopen(reopened);
	expect(shapeOf(grouped.sheets[0]?.drawings[0])).toMatchObject({
		kind: 'group',
		children: [{ name: 'Second', rotation: 90 }, { name: 'First' }],
	});
	const ungroupEdit = createEditSession(grouped);
	expect(ungroupEdit.ungroupDrawing(0, 0)).toEqual([0, 1]);
	const ungrouped = await reopen(grouped);
	expect(ungrouped.sheets[0]?.drawings.map((item) => shapeOf(item)?.geometry)).toEqual([
		'ellipse',
		'rect',
	]);
	expect(ungroupEdit.undo()).toBe(true);
	expect(grouped.sheets[0]?.drawings).toHaveLength(1);
});

it('round-trips a rotated picture and restores its orientation on undo', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	edit.addImage(
		0,
		Uint8Array.of(137, 80, 78, 71),
		'image/png',
		pictureAnchorAt({ row: 0, col: 0 }, 40, 20),
	);
	edit.rotateDrawing(0, 0, -90);
	expect((await reopen(book)).sheets[0]?.drawings[0]).toMatchObject({
		kind: 'image',
		rotation: 270,
	});
	expect(edit.undo()).toBe(true);
	expect((await reopen(book)).sheets[0]?.drawings[0]).not.toHaveProperty('rotation');
});

it('preserves visibility and flips on native and source-backed drawings', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	const anchor = pictureAnchorAt({ row: 0, col: 0 }, 70, 40);
	edit.addShape(0, anchor, { kind: 'preset', preset: 'rightArrow', name: 'Arrow' });
	edit.addImage(0, Uint8Array.of(137, 80, 78, 71), 'image/png', anchor);
	const loaded = await reopen(book);
	const change = createEditSession(loaded);
	change.flipDrawing(0, 0, 'horizontal');
	change.setDrawingHidden(0, 0, true);
	change.flipDrawing(0, 1, 'vertical');
	change.setDrawingHidden(0, 1, true);
	const hidden = await reopen(loaded);
	expect(shapeOf(hidden.sheets[0]?.drawings[0])).toMatchObject({ flipH: true, hidden: true });
	expect(hidden.sheets[0]?.drawings[1]).toMatchObject({ flipV: true, hidden: true });
	expect(change.undo()).toBe(true);
	expect((await reopen(loaded)).sheets[0]?.drawings[1]).not.toHaveProperty('hidden');
});
