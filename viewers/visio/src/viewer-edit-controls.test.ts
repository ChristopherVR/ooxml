import { afterEach, describe, expect, it, vi } from 'vitest';
import { ViewerController, type ViewerState } from './controller.js';
import { ViewerEditControls, editControlsTemplate } from './viewer-edit-controls.js';
import { demoDocument } from './demo-document.js';

function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.innerHTML = editControlsTemplate;
	const controller = new ViewerController();
	controller.setDocument(demoDocument);
	controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
	const controls = new ViewerEditControls(root, controller);
	const state = {
		...controller.state,
		edit: {
			sourceAvailable: true,
			busy: false,
			dirty: false,
			canUndo: false,
			canRedo: false,
			historyTruncated: false,
			error: null,
			diagnostics: [],
		},
	} as ViewerState;
	vi.spyOn(controller, 'state', 'get').mockImplementation(() => state);
	controls.render(state);
	const dispose = controls.wire();
	const button = (name: string) => root.querySelector<HTMLButtonElement>(`[data-edit="${name}"]`)!;
	const input = (value: string) => {
		controls.input.value = value;
		controls.input.dispatchEvent(new Event('input'));
	};
	return { root, controller, controls, state, button, input, dispose };
}
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});
describe('shared experimental text editing controls', () => {
	it('labels a plain-text target and forwards explicit page identity without interpreting text', async () => {
		const { root, controller, controls, button, input, dispose } = setup();
		const apply = vi.spyOn(controller, 'replacePlainText').mockResolvedValue();
		expect(controls.input.labels?.[0]?.textContent).toBe('Selected shape text');
		expect(root.querySelector('[data-edit-target]')?.textContent).toContain('Page ID: 1');
		input('<script>literal</script>\n+ - 0');
		button('apply').click();
		expect(apply).toHaveBeenCalledWith('1', 's1', '<script>literal</script>\n+ - 0');
		expect(root.querySelector('script')).toBeNull();
		await Promise.resolve();
		dispose();
	});
	it('retains textarea focus and cancels draft with Escape, without consuming native undo', () => {
		const { controls, input, button, dispose } = setup();
		const original = controls.input.value;
		controls.input.focus();
		input('draft');
		const undo = new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, cancelable: true });
		controls.input.dispatchEvent(undo);
		expect(undo.defaultPrevented).toBe(false);
		controls.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(controls.input.value).toBe(original);
		expect(button('apply').disabled).toBe(true);
		expect(controls.input.getRootNode()).toHaveProperty('activeElement', controls.input);
		dispose();
	});
	it('disables repeated commands while pending and renders safe errors', async () => {
		const { root, controller, controls, button, input, dispose } = setup();
		let reject!: (error: Error) => void;
		const apply = vi.spyOn(controller, 'replacePlainText').mockImplementation(
			() =>
				new Promise((_, fail) => {
					reject = fail;
				}),
		);
		input('draft');
		button('apply').click();
		button('apply').click();
		expect(apply).toHaveBeenCalledTimes(1);
		expect(controls.input.disabled).toBe(true);
		expect(button('undo').disabled).toBe(true);
		expect(button('cancel').disabled).toBe(false);
		reject(new Error('<img src=x> unsupported target'));
		await Promise.resolve();
		await Promise.resolve();
		expect(root.querySelector('[role="alert"]')?.textContent).toContain('<img src=x>');
		expect(root.querySelector('img')).toBeNull();
		expect(controls.input.value).toBe('draft');
		dispose();
	});
	it('clears drafts for target, page and document changes and respects model-only state', () => {
		const { controller, controls, state, input, button, dispose } = setup();
		input('draft');
		controls.render({ ...state, selectedShape: null });
		expect(controls.input.value).toBe('');
		expect(button('apply').disabled).toBe(true);
		controls.render(state);
		input('draft');
		controls.render({ ...state, pageIndex: 1 });
		expect(controls.input.value).not.toBe('draft');
		input('draft');
		controller.setDocument(demoDocument);
		controls.render(state);
		expect(controls.input.value).not.toBe('draft');
		controls.render({ ...state, edit: { ...state.edit, sourceAvailable: false } });
		expect(controls.input.disabled).toBe(true);
		expect(button('apply').disabled).toBe(true);
		dispose();
	});
	it('keeps background page IDs, history warnings and busy/loading availability', () => {
		const { controls, state, root, button, dispose } = setup();
		const background = { ...demoDocument.pages[0]!, id: 'background', isBackground: true };
		const document = { ...demoDocument, pages: [...demoDocument.pages, background] };
		controls.render({
			...state,
			document,
			selectedShape: { id: 's1', name: 'Background', pageId: 'background' },
			edit: { ...state.edit, dirty: true, canUndo: true, historyTruncated: true },
		});
		expect(root.querySelector('[data-edit-target]')?.textContent).toContain('Page ID: background');
		expect(root.querySelector('#edit-status')?.textContent).toContain('Earlier undo history');
		expect(button('undo').disabled).toBe(false);
		controls.render({
			...state,
			loading: true,
			edit: { ...state.edit, canUndo: true, canRedo: true },
		});
		expect(button('undo').disabled).toBe(true);
		expect(button('redo').disabled).toBe(true);
		dispose();
	});
});

it('shared controls drive a real core edit, refresh draft and restore byte-exact original on undo', async () => {
	const { editVsdx, parseVsdx } = await import('ooxml-core/visio');
	const { createVsdxFixture } = await import('../tests/fixture.mjs');
	const bytes = await createVsdxFixture('Original');
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		async (source, edits) => {
			const result = await editVsdx(source, edits);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
	);
	await controller.load(bytes);
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.innerHTML = editControlsTemplate;
	const controls = new ViewerEditControls(root, controller);
	const unsubscribe = controller.subscribe((state) => controls.render(state));
	const dispose = controls.wire();
	controller.selectShape({ id: '1', name: 'Import test', pageId: '1' });
	controls.input.value = '<script>safe & edited</script>';
	controls.input.dispatchEvent(new Event('input'));
	root.querySelector<HTMLButtonElement>('[data-edit="apply"]')!.click();
	await vi.waitFor(() => expect(controller.state.edit.dirty).toBe(true));
	expect(controls.input.value).toBe('<script>safe & edited</script>');
	expect(root.querySelector<HTMLButtonElement>('[data-edit="apply"]')!.disabled).toBe(true);
	expect(root.querySelector('script')).toBeNull();
	expect((await parseVsdx(controller.exportVsdx().bytes)).pages[0]!.shapes[0]!.text.plainText).toBe(
		'<script>safe & edited</script>',
	);
	root.querySelector<HTMLButtonElement>('[data-edit="undo"]')!.click();
	await vi.waitFor(() => expect(controller.state.edit.dirty).toBe(false));
	expect(controls.input.value).toBe('Original');
	expect(controller.exportVsdx().bytes).toEqual(Uint8Array.from(bytes));
	root.querySelector<HTMLButtonElement>('[data-edit="redo"]')!.click();
	await vi.waitFor(() => expect(controller.state.edit.dirty).toBe(true));
	expect(controls.input.value).toBe('<script>safe & edited</script>');
	dispose();
	unsubscribe();
	controller.destroy();
});

it('disposes safely after external controller destruction and ignores late edit rejection', async () => {
	const { controller, controls, input, button, dispose } = setup();
	let reject!: (error: Error) => void;
	vi.spyOn(controller, 'replacePlainText').mockImplementation(
		() =>
			new Promise((_, fail) => {
				reject = fail;
			}),
	);
	const original = controls.input.value;
	input('temporary draft');
	button('apply').click();
	controller.destroy();
	expect(() => dispose()).not.toThrow();
	expect(controls.input.value).toBe(original);
	reject(new Error('Obsolete operation'));
	await Promise.resolve();
	await Promise.resolve();
	expect(controls.input.value).toBe(original);
});

it('reconnects controls without retaining a draft or duplicating action listeners', async () => {
	const { controller, controls, state, input, button, dispose } = setup();
	const apply = vi.spyOn(controller, 'replacePlainText').mockResolvedValue();
	const original = controls.input.value;
	input('discarded draft');
	dispose();
	const disconnectAgain = controls.wire();
	controls.render(state);
	expect(controls.input.value).toBe(original);
	input('new draft');
	button('apply').click();
	expect(apply).toHaveBeenCalledTimes(1);
	await Promise.resolve();
	disconnectAgain();
});
