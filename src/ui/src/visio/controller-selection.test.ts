import { expect, it, vi } from 'vitest';
import { demoDocument, type VisioShapeSelection } from 'ooxml-core/visio/ui';
import JSZip from 'jszip';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { ViewerController } from './controller';

const selection = (id: string, pageId = '1'): VisioShapeSelection => ({ id, name: id, pageId });
function setup() {
	const model = structuredClone(demoDocument);
	const controller = new ViewerController();
	controller.setDocument(model);
	return { model, controller };
}

it('owns immutable ordered selection snapshots and keeps the first item primary', () => {
	const { controller } = setup();
	const events = vi.fn();
	controller.onEvent(events);
	const input = [{ ...selection('s1') }, { ...selection('s2') }];
	const initial = controller.state.selectedShapes;
	controller.selectShapes(input);
	const selected = controller.state.selectedShapes;
	input[0]!.name = 'Changed host name';
	input.pop();
	expect(selected.map((shape) => shape.id)).toEqual(['s1', 's2']);
	expect(selected[0]!.name).toBe('s1');
	expect(initial).toEqual([]);
	expect(Object.isFrozen(selected)).toBe(true);
	expect(selected.every(Object.isFrozen)).toBe(true);
	expect(controller.state.selectedShape).toBe(selected[0]);
	expect(events).toHaveBeenCalledWith('selection-change', selected);
	expect(events).toHaveBeenLastCalledWith('shape-select', selected[0]);
	controller.selectShape({ id: 's3', name: 'Single' });
	expect(controller.state.selectedShape).toEqual({ id: 's3', name: 'Single' });
	expect(controller.state.selectedShapes).toEqual([{ id: 's3', name: 'Single' }]);
	controller.selectShape(null);
	expect(controller.state.selectedShapes).toEqual([]);
	controller.destroy();
});

it('toggles membership without disturbing surviving order and deduplicates identities', () => {
	const { controller } = setup();
	controller.selectShapes([selection('s1'), selection('s2'), selection('s1')]);
	expect(controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['s1', 's2']);
	controller.toggleShapeSelection(selection('s1'));
	expect(controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['s2']);
	controller.toggleShapeSelection(selection('s3'));
	expect(controller.state.selectedShape?.id).toBe('s2');
	controller.toggleShapeSelection(selection('s3'));
	expect(controller.state.selectedShape?.id).toBe('s2');
	controller.clearSelection();
	expect(controller.state.selectedShape).toBeNull();
	controller.destroy();
});

it('normalizes parent and descendant selection to avoid applying transforms twice', () => {
	const { model, controller } = setup();
	const group = model.pages[0]!.shapes[0]!;
	group.kind = 'group';
	group.children = [structuredClone(model.pages[0]!.shapes[2]!)];
	model.pages[0]!.shapes.splice(2, 1);
	controller.setDocument(model);
	controller.selectShapes([selection('s2'), selection('s1')]);
	expect(controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['s1']);
	controller.selectShapes([selection('s1'), selection('s2')]);
	expect(controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['s1']);
	controller.selectShapes([selection('s2')]);
	expect(controller.state.selectedShape?.id).toBe('s2');
	controller.selectAll();
	expect(controller.state.selectedShapes.map((shape) => shape.id)).not.toContain('s2');
	controller.destroy();
});

it('keeps page-scoped identities distinct and selects all only on the active page', () => {
	const { model, controller } = setup();
	const front = model.pages[0]!,
		background = model.pages[1]!;
	front.backgroundPageId = background.id;
	background.isBackground = true;
	background.shapes = [structuredClone(front.shapes[0]!)];
	controller.setDocument(model);
	controller.selectShapes([selection('s1', '1'), selection('s1', '2')]);
	expect(controller.state.selectedShapes).toHaveLength(2);
	controller.selectAll();
	expect(controller.state.selectedShapes.every((shape) => shape.pageId === '1')).toBe(true);
	expect(controller.state.selectedShapes).toHaveLength(front.shapes.length);
	controller.setPage(1);
	expect(controller.state.selectedShapes).toEqual([]);
	controller.selectAll();
	expect(controller.state.selectedShapes).toEqual(
		[selection('s1', '2')].map((item) => ({ ...item, name: background.shapes[0]!.name })),
	);
	controller.setDocument(model);
	expect(controller.state.selectedShapes).toEqual([]);
	controller.destroy();
});

it('prunes hidden layer targets while preserving the rest of the selection', () => {
	const { model, controller } = setup();
	model.pages[0]!.layers = [
		{ id: '0', name: 'Layer', visible: true, printable: true, locked: false },
	];
	model.pages[0]!.shapes[0]!.layerIds = ['0'];
	controller.setDocument(model);
	controller.selectShapes([selection('s2'), selection('s1')]);
	controller.setLayerVisibility('1', '0', false);
	expect(controller.state.selectedShapes).toEqual([selection('s2')]);
	expect(controller.state.selectedShape?.id).toBe('s2');
	controller.resetLayerVisibility();
	controller.selectShapes([selection('s1'), selection('missing')]);
	expect(controller.state.selectedShapes).toEqual([selection('s1')]);
	controller.destroy();
});

it('prunes deleted targets on edit/history acceptance and preserves surviving selection', async () => {
	const original = structuredClone(demoDocument);
	const changed = structuredClone(original);
	changed.pages[0]!.shapes = changed.pages[0]!.shapes.filter((shape) => shape.id !== 's2');
	const parser = async (bytes: Uint8Array | ArrayBuffer) =>
		new Uint8Array(bytes instanceof Uint8Array ? bytes : bytes)[0] === 1 ? original : changed;
	const controller = new ViewerController(
		parser,
		() => {},
		async () => ({
			bytes: new Uint8Array([2]),
			document: changed,
			changedParts: ['page.xml'],
			diagnostics: [],
		}),
	);
	await controller.load(new Uint8Array([1]));
	controller.selectShapes([selection('s1'), selection('s2')]);
	await controller.applyEdits([{ type: 'delete-shape', pageId: '1', shapeId: 's2' }]);
	expect(controller.state.selectedShapes).toEqual([selection('s1')]);
	await controller.undo();
	expect(controller.state.selectedShapes).toEqual([selection('s1'), selection('s2')]);
	expect(controller.state.selectedShape).toBe(controller.state.selectedShapes[0]);
	expect(Object.isFrozen(controller.state.selectedShapes)).toBe(true);
	await controller.redo();
	expect(controller.state.selectedShapes).toEqual([selection('s1')]);
	controller.destroy();
	expect(controller.state.selectedShapes).toEqual([]);
	expect(() => controller.selectAll()).toThrow('destroyed');
	expect(() => controller.selectShapes([])).toThrow('destroyed');
});

it('rejects invalid host selections before changing state and suppresses stale reentrant events', () => {
	const { controller } = setup();
	const state = controller.state;
	expect(() =>
		controller.selectShapes([{ id: 's1', name: 1 } as unknown as VisioShapeSelection]),
	).toThrow('Selection requires');
	expect(controller.state).toBe(state);
	const seen: string[] = [];
	controller.onEvent((name, detail) => {
		if (
			name === 'selection-change' &&
			(detail as readonly VisioShapeSelection[]).at(-1)?.id === 's1'
		)
			controller.selectShape(selection('s2'));
	});
	controller.onEvent((name, detail) => {
		if (name === 'selection-change')
			seen.push((detail as readonly VisioShapeSelection[]).at(-1)?.id ?? 'none');
	});
	controller.selectShape(selection('s1'));
	expect(seen).toEqual(['s2']);
	expect(controller.state.selectedShape?.id).toBe('s2');
	controller.destroy();
});

it('keeps every source-backed formatting target selected through atomic edits and history', async () => {
	const zip = await JSZip.loadAsync(await createVsdxFixture('First'));
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	const original = xml.match(/<Shape ID="1"[\s\S]*?<\/Shape>/)![0];
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(
			'</Shapes>',
			`${original.replace('ID="1"', 'ID="2"').replace('First', 'Second')}</Shapes>`,
		),
	);
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		async (bytes, edits) => {
			const result = await editVsdx(bytes, edits);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
	);
	await controller.load(await zip.generateAsync({ type: 'uint8array' }));
	const page = controller.state.document!.pages[0]!;
	controller.selectShapes(
		page.shapes.map((shape) => ({ id: shape.id, name: shape.name, pageId: page.id })),
	);
	const before = controller.state.selectedShapes;
	const selectionChanged = vi.fn();
	controller.onEvent((name) => {
		if (name === 'selection-change') selectionChanged();
	});
	await controller.applyEdits(
		before.map((shape) => ({
			type: 'format-text',
			pageId: page.id,
			shapeId: shape.id,
			bold: true,
		})),
	);
	expect(controller.state.selectedShapes).toBe(before);
	expect(selectionChanged).not.toHaveBeenCalled();
	expect(
		controller.state.document!.pages[0]!.shapes.every((shape) => shape.text.runs![0]!.bold),
	).toBe(true);
	await controller.undo();
	expect(controller.state.selectedShapes).toBe(before);
	await controller.redo();
	expect(controller.state.selectedShapes).toBe(before);
	const reopened = await parseVsdx(controller.exportVsdx().bytes);
	expect(reopened.pages[0]!.shapes.every((shape) => shape.text.runs![0]!.bold)).toBe(true);
	controller.destroy();
});

it.each(['many', 'single', 'toggle'] as const)(
	'does not let reentrant host getters overwrite newer state through %s selection',
	(method) => {
		for (const effect of ['replace', 'destroy', 'select', 'page'] as const) {
			const { controller } = setup();
			const events = vi.fn();
			controller.onEvent(events);
			let invoked = false;
			const reenter = () => {
				if (invoked) return;
				invoked = true;
				if (effect === 'replace') controller.setDocument(null);
				else if (effect === 'destroy') controller.destroy();
				else if (effect === 'select') controller.selectShapes([selection('s2')]);
				else controller.setPage(1);
			};
			const target = {
				get id() {
					if (method !== 'many') reenter();
					return 's1';
				},
				get name() {
					if (method === 'many') reenter();
					return 'Stale target';
				},
				pageId: '1',
			};
			if (method === 'many') controller.selectShapes([target]);
			else if (method === 'single') controller.selectShape(target);
			else controller.toggleShapeSelection(target);
			expect(invoked, `${method}/${effect}`).toBe(true);
			expect(controller.state.selectedShapes).toEqual(effect === 'select' ? [selection('s2')] : []);
			if (effect === 'replace' || effect === 'destroy')
				expect(controller.state.document).toBeNull();
			if (effect === 'page') expect(controller.state.pageIndex).toBe(1);
			for (const [name, detail] of events.mock.calls) {
				if (name === 'selection-change') expect(detail).not.toContainEqual(target);
				if (name === 'shape-select') expect(detail?.id).not.toBe('s1');
			}
			controller.destroy();
		}
	},
);
