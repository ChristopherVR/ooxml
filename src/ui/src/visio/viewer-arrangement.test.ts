import { afterEach, expect, it } from 'vitest';
import JSZip from 'jszip';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';

afterEach(() => document.body.replaceChildren());

it('selects all and clears selection from ribbon and canvas keys while keeping text-field shortcuts', async () => {
	const ui = await setup(false);
	const key = (name: string, control = false, target: Element = ui.viewport) =>
		target.dispatchEvent(
			new KeyboardEvent('keydown', {
				key: name,
				ctrlKey: control,
				bubbles: true,
				composed: true,
				cancelable: true,
			}),
		);
	ui.press('select-all');
	expect(ui.controller.state.selectedShapes).toHaveLength(1);
	ui.press('clear-selection');
	expect(ui.controller.state.selectedShapes).toHaveLength(0);
	expect(key('a', true)).toBe(false);
	expect(ui.controller.state.selectedShapes).toHaveLength(1);
	expect(key('Escape')).toBe(false);
	expect(ui.controller.state.selectedShapes).toHaveLength(0);
	const textarea = document.createElement('textarea');
	ui.viewport.append(textarea);
	expect(key('a', true, textarea)).toBe(true);
	expect(ui.controller.state.selectedShapes).toHaveLength(0);
	ui.dispose();
	ui.controller.destroy();
});

it.each([
	['left', 3, 2],
	['center', 4, 2],
	['right', 5, 2],
	['top', 2, 6],
	['middle', 2, 7],
	['bottom', 2, 8],
] as const)(
	'aligns %s with one source batch and restores source and selection on undo',
	async (edge, x, y) => {
		const ui = await setup();
		await ui.controller.applyEdits([
			{ type: 'create-rectangle', pageId: '1', shapeId: '2', x: 2, y: 2, width: 1, height: 3 },
		]);
		const original = ui.controller.exportVsdx().bytes;
		ui.controller.selectShapes([
			{ id: '1', name: 'Anchor', pageId: '1' },
			{ id: '2', name: 'Moved', pageId: '1' },
		]);
		expect(ui.button(`align-shapes-${edge}`).disabled).toBe(false);
		ui.press(`align-shapes-${edge}`);
		await ui.done();
		expect(ui.controller.state.document!.pages[0]!.shapes[1]!.rotation).toMatchObject({
			pinX: x,
			pinY: y,
			angle: 0,
		});
		expect(ui.shape().rotation).toMatchObject({ pinX: 4, pinY: 7 });
		expect(ui.edits.at(-1)).toHaveLength(1);
		ui.press('undo');
		await ui.done();
		expect(ui.controller.exportVsdx().bytes).toEqual(original);
		expect(ui.controller.state.selectedShapes.map((item) => item.id)).toEqual(['1', '2']);
		ui.dispose();
		ui.controller.destroy();
	},
);

it('distributes three shapes with equal edge gaps and refuses independent multi-rotation', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{ type: 'create-rectangle', pageId: '1', shapeId: '2', x: 2, y: 2, width: 1, height: 3 },
		{ type: 'create-rectangle', pageId: '1', shapeId: '3', x: 9, y: 10, width: 2, height: 1 },
	]);
	ui.press('select-all');
	expect(ui.button('rotate-left').disabled).toBe(true);
	const count = ui.edits.length;
	ui.commands.run({ type: 'rotate', direction: 'left' });
	expect(ui.edits).toHaveLength(count);
	ui.press('distribute-horizontal');
	await ui.done();
	const byId = new Map(
		ui.controller.state.document!.pages[0]!.shapes.map((item) => [item.id, item]),
	);
	expect(byId.get('2')!.rotation!.pinX).toBe(2);
	expect(byId.get('3')!.rotation!.pinX).toBe(9);
	expect(byId.get('1')!.rotation!.pinX).toBeCloseTo(5.25);
	ui.press('distribute-vertical');
	await ui.done();
	expect(ui.shape().rotation!.pinY).toBeCloseTo(6.5);
	ui.dispose();
	ui.controller.destroy();
});

it('deletes all selected source targets atomically and preserves bytes when a later target refuses', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{ type: 'create-rectangle', pageId: '1', shapeId: '2', x: 2, y: 2, width: 1, height: 3 },
	]);
	const before = ui.controller.exportVsdx().bytes;
	ui.press('select-all');
	ui.viewport.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));
	await ui.done();
	expect(ui.controller.state.document!.pages[0]!.shapes).toHaveLength(0);
	expect(ui.controller.state.selectedShapes).toHaveLength(0);
	ui.press('undo');
	await ui.done();
	expect(ui.controller.exportVsdx().bytes).toEqual(before);
	const zip = await JSZip.loadAsync(before);
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(/<Shape\b[^>]*\bID="2"[^>]*>/u, '$&<Cell N="LockDelete" V="1"/>'),
	);
	const locked = await zip.generateAsync({ type: 'uint8array' });
	await ui.controller.load(locked);
	ui.press('select-all');
	ui.commands.run({ type: 'delete' });
	await ui.done();
	expect(ui.controller.state.edit.error).toBeDefined();
	expect(ui.controller.exportVsdx().bytes).toEqual(locked);
	expect(ui.controller.state.document!.pages[0]!.shapes).toHaveLength(2);
	expect(ui.controller.state.selectedShapes).toHaveLength(2);
	ui.dispose();
	ui.controller.destroy();
});

it('preserves the complete source and selection when a later alignment target is locked', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{ type: 'create-rectangle', pageId: '1', shapeId: '2', x: 2, y: 2, width: 1, height: 3 },
		{ type: 'create-rectangle', pageId: '1', shapeId: '3', x: 9, y: 10, width: 2, height: 1 },
	]);
	const zip = await JSZip.loadAsync(ui.controller.exportVsdx().bytes);
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(/<Shape\b[^>]*\bID="3"[^>]*>/u, '$&<Cell N="LockMoveX" V="1"/>'),
	);
	const locked = await zip.generateAsync({ type: 'uint8array' });
	await ui.controller.load(locked);
	ui.press('select-all');
	const selection = ui.controller.state.selectedShapes;
	ui.press('align-shapes-left');
	await ui.done();
	expect(ui.edits.at(-1)).toHaveLength(2);
	expect(ui.controller.state.edit.error).toBeDefined();
	expect(ui.controller.exportVsdx().bytes).toEqual(locked);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	expect(ui.controller.state.selectedShapes).toBe(selection);
	ui.dispose();
	ui.controller.destroy();
});
