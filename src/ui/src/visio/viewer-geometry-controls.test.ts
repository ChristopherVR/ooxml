import { afterEach, expect, it, vi } from 'vitest';
import { ViewerController } from './controller';
import { ViewerEditControls, editControlsTemplate } from './viewer-edit-controls';
import { demoDocument } from 'ooxml-core/visio/ui';

async function setup(line = false) {
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.innerHTML = editControlsTemplate;
	const controller = new ViewerController(async () => {
		const model = structuredClone(demoDocument);
		if (line) {
			model.pages[0]!.shapes[0]!.kind = 'connector';
			model.pages[0]!.shapes[0]!.height = 0;
		}
		return model;
	});
	await controller.load(new Uint8Array([1]));
	controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
	const controls = new ViewerEditControls(root, controller);
	const unsubscribe = controller.subscribe((state) => controls.render(state));
	const dispose = controls.wire();
	const input = (name: string, value: string) => {
		const field = root.querySelector<HTMLInputElement>(`[data-geometry-field="${name}"]`)!;
		field.value = value;
		field.dispatchEvent(new Event('input'));
	};
	const button = (action: string) =>
		root.querySelector<HTMLButtonElement>(`[data-geometry-action="${action}"]`)!;
	return {
		root,
		controller,
		input,
		button,
		dispose: () => {
			unsubscribe();
			dispose();
			controller.destroy();
		},
	};
}
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});
it('requires explicit pin coordinates and IDs and forwards all geometry actions', async () => {
	const { root, controller, input, button, dispose } = await setup();
	const apply = vi.spyOn(controller, 'applyEdits').mockResolvedValue();
	expect(root.textContent).toContain('bottom-left origin, Y up');
	expect(button('move-shape').disabled).toBe(true);
	input('x', '-1.25');
	input('y', '3.5');
	input('width', '2');
	input('height', '4');
	button('move-shape').click();
	expect(apply).toHaveBeenLastCalledWith([
		{ type: 'move-shape', pageId: '1', shapeId: 's1', x: -1.25, y: 3.5 },
	]);
	button('resize-shape').click();
	expect(apply).toHaveBeenLastCalledWith([
		{ type: 'resize-shape', pageId: '1', shapeId: 's1', width: 2, height: 4 },
	]);
	expect(button('create-rectangle').disabled).toBe(true);
	input('id', '42');
	button('create-rectangle').click();
	expect(apply).toHaveBeenLastCalledWith([
		{ type: 'create-rectangle', pageId: '1', shapeId: '42', x: -1.25, y: 3.5, width: 2, height: 4 },
	]);
	button('delete-shape').click();
	expect(apply).toHaveBeenLastCalledWith([{ type: 'delete-shape', pageId: '1', shapeId: 's1' }]);
	dispose();
});
it('rejects incomplete dimensions, resets drafts on selection changes and permits create without selection', async () => {
	const { controller, input, button, dispose } = await setup();
	input('width', '0');
	input('height', '2');
	expect(button('resize-shape').disabled).toBe(true);
	input('x', '2');
	input('y', '3');
	controller.selectShape(null);
	expect(button('delete-shape').disabled).toBe(true);
	expect(button('move-shape').disabled).toBe(true);
	input('id', '42');
	input('x', '2');
	input('y', '3');
	input('width', '1');
	input('height', '2');
	expect(button('create-rectangle').disabled).toBe(false);
	dispose();
});
it('admits zero Height only for selected line width controls while rectangle creation stays positive', async () => {
	const { controller, input, button, dispose } = await setup(true);
	const apply = vi.spyOn(controller, 'applyEdits').mockResolvedValue();
	input('width', '4');
	input('height', '0');
	expect(button('resize-shape').disabled).toBe(false);
	button('resize-shape').click();
	expect(apply).toHaveBeenCalledWith([
		{ type: 'resize-shape', pageId: '1', shapeId: 's1', width: 4, height: 0 },
	]);
	input('id', '42');
	input('x', '2');
	input('y', '3');
	expect(button('create-rectangle').disabled).toBe(true);
	dispose();
});
it('shows safe core rejection text and forgets obsolete errors on navigation', async () => {
	const { root, controller, button, dispose } = await setup();
	vi.spyOn(controller, 'applyEdits').mockRejectedValue(
		new Error('<script>Referenced deletion rejected</script>'),
	);
	button('delete-shape').click();
	await Promise.resolve();
	await Promise.resolve();
	const error = root.querySelector<HTMLElement>('[data-geometry-error]')!;
	expect(error.hidden).toBe(false);
	expect(error.textContent).toContain('Referenced deletion rejected');
	expect(error.querySelector('script')).toBeNull();
	controller.selectShape(null);
	expect(error.hidden).toBe(true);
	dispose();
});

it('forwards optional rectangle text literally and retains rejected geometry drafts and codes', async () => {
	const { root, controller, input, button, dispose } = await setup();
	const apply = vi
		.spyOn(controller, 'applyEdits')
		.mockRejectedValue(
			Object.assign(new Error('ID already exists'), { code: 'EDIT_DUPLICATE_SHAPE' }),
		);
	input('id', '42');
	input('x', '2');
	input('y', '3');
	input('width', '1');
	input('height', '2');
	input('text', '<script>Literal rectangle</script>');
	button('create-rectangle').click();
	await Promise.resolve();
	await Promise.resolve();
	expect(apply).toHaveBeenCalledWith([
		{
			type: 'create-rectangle',
			pageId: '1',
			shapeId: '42',
			x: 2,
			y: 3,
			width: 1,
			height: 2,
			text: '<script>Literal rectangle</script>',
		},
	]);
	expect(root.querySelector<HTMLInputElement>('[data-geometry-field="id"]')!.value).toBe('42');
	expect(root.querySelector('[data-geometry-error]')?.textContent).toBe(
		'EDIT_DUPLICATE_SHAPE: ID already exists',
	);
	expect(root.querySelector('script')).toBeNull();
	dispose();
});

it('source replacement cancels the geometry draft identity and ignores an obsolete worker refusal', async () => {
	const { root, controller, input, button, dispose } = await setup();
	let reject!: (cause: Error) => void;
	vi.spyOn(controller, 'applyEdits').mockImplementation(
		() =>
			new Promise((_, fail) => {
				reject = fail;
			}),
	);
	input('x', '2');
	input('y', '3');
	button('move-shape').click();
	await controller.load(new Uint8Array([3]));
	expect(root.querySelector<HTMLInputElement>('[data-geometry-field="x"]')!.value).toBe('');
	reject(Object.assign(new Error('Obsolete protected target'), { code: 'EDIT_PROTECTED_CELL' }));
	await Promise.resolve();
	await Promise.resolve();
	expect(root.querySelector<HTMLElement>('[data-geometry-error]')!.hidden).toBe(true);
	expect(controller.state.edit.dirty).toBe(false);
	dispose();
});

it('forwards angular drafts in radians and clears them with shared selection state', async () => {
	const { controller, input, button, dispose } = await setup();
	const apply = vi.spyOn(controller, 'applyEdits').mockResolvedValue();
	expect(button('rotate-shape').disabled).toBe(true);
	input('angle', '-45');
	expect(button('rotate-shape').disabled).toBe(false);
	button('rotate-shape').click();
	expect(apply).toHaveBeenLastCalledWith([
		{ type: 'rotate-shape', pageId: '1', shapeId: 's1', angle: -Math.PI / 4 },
	]);
	controller.selectShape(null);
	expect(button('rotate-shape').disabled).toBe(true);
	dispose();
});
it('keeps unsupported connector rotation disabled', async () => {
	const { input, button, dispose } = await setup(true);
	input('angle', '30');
	expect(button('rotate-shape').disabled).toBe(true);
	dispose();
});
