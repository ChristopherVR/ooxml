import { afterEach, expect, it } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { setupFormattingViewer } from './__fixtures__/formatting-viewer';
import { createContextMenus } from './viewer-context-menu';
import { pointerShapeTarget } from './viewer-shape-target';

afterEach(() => document.body.replaceChildren());
const setup = async (source = true) => {
	const view = await setupFormattingViewer(source);
	view.root.append(...createContextMenus(document));
	view.commands.render(view.controller.state);
	if (source)
		await view.controller.applyEdits([
			{ type: 'create-rectangle', pageId: '1', shapeId: '2', x: 2, y: 2, width: 1, height: 1 },
			{ type: 'create-ellipse', pageId: '1', shapeId: '3', x: 4, y: 3, width: 2, height: 1 },
		]);
	return view;
};
const select = (view: Awaited<ReturnType<typeof setup>>, ids: string[]) =>
	view.controller.selectShapes(ids.map((id) => ({ id, name: '', pageId: '1' })));
const key = (letter: string) =>
	new KeyboardEvent('keydown', {
		key: letter,
		ctrlKey: true,
		shiftKey: true,
		bubbles: true,
		composed: true,
		cancelable: true,
	});

it('groups the selection, selects the group and ungroups with undo and redo', async () => {
	const view = await setup();
	const before = view.controller.state.document!.pages[0]!;
	select(view, ['2', '3']);
	expect(view.button('group-shapes').disabled).toBe(false);
	expect(view.button('ungroup').disabled).toBe(true);
	view.press('group-shapes');
	await view.done();
	const grouping = view.edits.at(-1)![0]!;
	expect(grouping).toEqual({
		type: 'group-shapes',
		pageId: '1',
		shapeId: '4',
		memberIds: ['2', '3'],
	});
	let page = view.controller.state.document!.pages[0]!;
	expect(page.shapes.map((shape) => shape.id)).toEqual(['1', '4']);
	expect(page.shapes[1]!.children.map((shape) => shape.id)).toEqual(['2', '3']);
	expect(view.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['4']);
	expect(view.button('group-shapes').disabled).toBe(true);
	expect(view.button('ungroup').disabled).toBe(false);
	expect(view.button('ctx-ungroup').disabled).toBe(false);

	const ungroup = key('u');
	view.viewport.dispatchEvent(ungroup);
	expect(ungroup.defaultPrevented).toBe(true);
	await view.done();
	page = view.controller.state.document!.pages[0]!;
	expect(page.shapes.map((shape) => shape.id)).toEqual(['1', '2', '3']);
	for (const [index, shape] of page.shapes.entries())
		for (let i = 0; i < 6; i++)
			expect(shape.transform[i]).toBeCloseTo(before.shapes[index]!.transform[i]!, 9);
	expect(view.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['2', '3']);
	expect((await parseVsdx(view.controller.exportVsdx().bytes)).pages[0]!.shapes).toHaveLength(3);

	await view.controller.undo();
	expect(view.controller.state.document!.pages[0]!.shapes.map((shape) => shape.id)).toEqual([
		'1',
		'4',
	]);
	expect(view.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['4']);
	await view.controller.redo();
	expect(view.controller.state.document!.pages[0]!.shapes).toHaveLength(3);
	view.dispose();
	view.controller.destroy();
});

it('groups from Ctrl+Shift+G and the context menu, and moves the group as one shape', async () => {
	const view = await setup();
	select(view, ['2', '3']);
	const group = key('g');
	view.viewport.dispatchEvent(group);
	await view.done();
	expect(view.edits.at(-1)![0]!.type).toBe('group-shapes');
	const created = view.controller.state.document!.pages[0]!.shapes[1]!;
	await view.controller.applySelectionEdits([
		{
			type: 'move-shape',
			pageId: '1',
			shapeId: created.id,
			x: created.rotation!.pinX + 1,
			y: created.rotation!.pinY,
		},
	]);
	const moved = view.controller.state.document!.pages[0]!.shapes[1]!;
	expect(moved.transform[4]).toBeCloseTo(created.transform[4] + 1, 9);
	expect(moved.children).toEqual(created.children);
	await view.controller.undo();
	await view.controller.undo();
	select(view, ['2', '3']);
	view.press('ctx-group');
	await view.done();
	expect(view.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['4']);
	view.dispose();
	view.controller.destroy();
});

it('disables Group and Ungroup for read-only, single, busy and inadmissible selections', async () => {
	const readOnly = await setup(false);
	expect(readOnly.button('group-shapes').disabled).toBe(true);
	expect(readOnly.button('ctx-ungroup').disabled).toBe(true);
	readOnly.commands.run({ type: 'grouping', operation: 'group' });
	expect(readOnly.edits).toEqual([]);
	readOnly.dispose();
	readOnly.controller.destroy();

	const view = await setup();
	const edits = view.edits.length;
	select(view, ['2']);
	expect(view.button('group-shapes').disabled).toBe(true);
	expect(view.button('ungroup').disabled).toBe(true);
	view.commands.run({ type: 'grouping', operation: 'ungroup' });
	select(view, ['2', '3']);
	view.commands.render({
		...view.controller.state,
		edit: { ...view.controller.state.edit, busy: true },
	});
	expect(view.button('group-shapes').disabled).toBe(true);
	view.commands.render(view.controller.state);
	expect(view.button('group-shapes').disabled).toBe(false);
	await view.controller.applyEdits([
		{ type: 'create-line', pageId: '1', shapeId: '5', beginX: 1, beginY: 1, endX: 2, endY: 2 },
	]);
	select(view, ['2', '5']);
	expect(view.button('group-shapes').disabled).toBe(true);
	expect(view.edits).toHaveLength(edits + 1);
	view.dispose();
	view.controller.destroy();
});

it('targets the outermost group first and subselects members on a second click', () => {
	const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
	const g = (id: string) => {
		const node = document.createElementNS('http://www.w3.org/2000/svg', 'g');
		node.dataset.shapeId = id;
		return node;
	};
	const outer = g('4'),
		middle = g('5'),
		leaf = g('2');
	const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
	leaf.append(path);
	middle.append(leaf);
	outer.append(middle);
	svg.append(outer);
	expect(pointerShapeTarget(path, [])).toBe(outer);
	expect(pointerShapeTarget(path, [{ id: '4' }])).toBe(middle);
	expect(pointerShapeTarget(path, [{ id: '5' }])).toBe(leaf);
	expect(pointerShapeTarget(path, [{ id: '2' }])).toBe(leaf);
	expect(pointerShapeTarget(path, [{ id: '4' }], true)).toBe(outer);
	expect(pointerShapeTarget(svg, [])).toBeNull();
});
