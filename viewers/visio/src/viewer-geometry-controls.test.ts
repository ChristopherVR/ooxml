import { afterEach, expect, it, vi } from 'vitest';
import { ViewerController } from './controller.js';
import { ViewerEditControls, editControlsTemplate } from './viewer-edit-controls.js';
import { demoDocument } from './demo-document.js';

async function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.innerHTML = editControlsTemplate;
	const controller = new ViewerController(async () => structuredClone(demoDocument));
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
