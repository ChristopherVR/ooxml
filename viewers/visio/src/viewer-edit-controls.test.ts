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
	const geometryDraft = root.querySelector<HTMLInputElement>('[data-geometry-field="width"]')!;
	geometryDraft.value = '4';
	geometryDraft.dispatchEvent(new Event('input'));
	controls.input.value = '<script>safe & edited</script>';
	controls.input.dispatchEvent(new Event('input'));
	root.querySelector<HTMLButtonElement>('[data-edit="apply"]')!.click();
	await vi.waitFor(() => expect(controller.state.edit.dirty).toBe(true));
	expect(controls.input.value).toBe('<script>safe & edited</script>');
	expect(geometryDraft.value).toBe('4');
	expect(root.querySelector<HTMLButtonElement>('[data-edit="apply"]')!.disabled).toBe(true);
	expect(root.querySelector('script')).toBeNull();
	expect((await parseVsdx(controller.exportVsdx().bytes)).pages[0]!.shapes[0]!.text.plainText).toBe(
		'<script>safe & edited</script>',
	);
	root.querySelector<HTMLButtonElement>('[data-edit="undo"]')!.click();
	await vi.waitFor(() => expect(controller.state.edit.dirty).toBe(false));
	expect(controls.input.value).toBe('Original');
	expect(geometryDraft.value).toBe('4');
	expect(controller.exportVsdx().bytes).toEqual(Uint8Array.from(bytes));
	root.querySelector<HTMLButtonElement>('[data-edit="redo"]')!.click();
	await vi.waitFor(() => expect(controller.state.edit.dirty).toBe(true));
	expect(controls.input.value).toBe('<script>safe & edited</script>');
	expect(geometryDraft.value).toBe('4');
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

it('preserves unapplied text across same-source geometry/history updates and resets on source replacement', async () => {
	const { controller, controls, state, input, dispose } = setup();
	input('Unapplied text');
	const updated = structuredClone(state.document!);
	updated.pages[0]!.shapes[0]!.width = 4;
	controls.render({ ...state, document: updated });
	expect(controls.input.value).toBe('Unapplied text');
	updated.pages[0]!.shapes[0]!.text.plainText = 'Restored source text';
	controls.render({ ...state, document: updated });
	expect(controls.input.value).toBe('Unapplied text');
	controller.setDocument(updated);
	controls.render({ ...state, document: updated });
	expect(controls.input.value).toBe('Restored source text');
	dispose();
});
it('Cancel handles geometry-only drafts and clears every field', () => {
	const { root, controller, controls, state, button, dispose } = setup();
	const cancel = vi.spyOn(controller, 'cancelEdit');
	const x = root.querySelector<HTMLInputElement>('[data-geometry-field="x"]')!;
	x.value = '4';
	x.dispatchEvent(new Event('input'));
	expect(button('cancel').disabled).toBe(false);
	button('cancel').click();
	expect(cancel).toHaveBeenCalledOnce();
	expect(x.value).toBe('');
	expect(button('cancel').disabled).toBe(true);
	expect(controls.input.value).toBe(state.document!.pages[0]!.shapes[0]!.text.plainText);
	dispose();
});
it('shows the core refusal code as plain text while retaining the text draft', async () => {
	const { root, controller, controls, button, input, dispose } = setup();
	vi.spyOn(controller, 'replacePlainText').mockRejectedValue(
		Object.assign(new Error('Protected text'), { code: 'EDIT_PROTECTED_CELL' }),
	);
	input('Keep this draft');
	button('apply').click();
	await Promise.resolve();
	await Promise.resolve();
	expect(root.querySelector('[data-edit-error]')?.textContent).toBe(
		'EDIT_PROTECTED_CELL: Protected text',
	);
	expect(controls.input.value).toBe('Keep this draft');
	dispose();
});

it.each(['text', 'geometry'] as const)(
	'retains %s draft without obsolete AbortError while a replacement load is cancelled or fails',
	async (action) => {
		for (const replacement of ['cancel', 'fail'] as const) {
			let rejectEdit: ((cause: Error) => void) | undefined;
			const editor = Object.assign(
				() =>
					new Promise<never>((_, reject) => {
						rejectEdit = reject;
					}),
				{ cancel: () => rejectEdit?.(new DOMException('Superseded edit', 'AbortError')) },
			);
			const controller = new ViewerController(
				async () => structuredClone(demoDocument),
				() => {},
				editor,
			);
			await controller.load(new Uint8Array([1]));
			controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
			const source = controller.sourceGeneration;
			const host = document.createElement('div');
			document.body.append(host);
			const root = host.attachShadow({ mode: 'open' });
			root.innerHTML = editControlsTemplate;
			const controls = new ViewerEditControls(root, controller);
			const unsubscribe = controller.subscribe((state) => controls.render(state)),
				dispose = controls.wire();
			controls.input.value = 'Keep pending text';
			controls.input.dispatchEvent(new Event('input'));
			const field = (name: string) =>
				root.querySelector<HTMLInputElement>(`[data-geometry-field="${name}"]`)!;
			if (action === 'geometry') {
				field('x').value = '2';
				field('y').value = '3';
				field('y').dispatchEvent(new Event('input'));
			}
			root
				.querySelector<HTMLButtonElement>(
					action === 'text' ? '[data-edit="apply"]' : '[data-geometry-action="move-shape"]',
				)!
				.click();
			let rejectRead!: (cause: Error) => void;
			const loading = controller
				.loadSource(
					() =>
						new Promise<never>((_, reject) => {
							rejectRead = reject;
						}),
				)
				.catch((error: unknown) => error);
			await Promise.resolve();
			await Promise.resolve();
			await Promise.resolve();
			if (replacement === 'cancel') controller.cancelLoad();
			rejectRead(new Error('Replacement read failed'));
			await loading;
			await vi.waitFor(() => expect(controller.state.edit.busy).toBe(false));
			expect(controller.sourceGeneration).toBe(source);
			expect(controller.exportVsdx().bytes).toEqual(new Uint8Array([1]));
			expect(controls.input.value).toBe('Keep pending text');
			if (action === 'geometry') expect(field('x').value).toBe('2');
			expect(root.querySelector<HTMLElement>('[data-edit-error]')!.hidden).toBe(true);
			expect(root.querySelector<HTMLElement>('[data-geometry-error]')!.hidden).toBe(true);
			dispose();
			unsubscribe();
			controller.destroy();
			host.remove();
		}
	},
);
