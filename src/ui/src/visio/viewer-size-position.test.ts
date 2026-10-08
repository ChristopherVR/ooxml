import { afterEach, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import {
	captureVisioClipboard,
	createVsdx,
	editVsdx,
	parseVsdx,
	serializeVisioClipboard,
} from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { createSizePosition, ViewerSizePosition } from './viewer-size-position';
import type { CancellableEditor } from './worker-editor';
import { mountViewer } from './binding';
import { emitRibbonAction } from './ribbon-action';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});
const edit: CancellableEditor = async (bytes, commands) => {
	const result = await editVsdx(bytes, commands);
	return { ...result, document: await parseVsdx(result.bytes) };
};
async function fixture(scale = 1, lock = false): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await createVsdx());
	const path = 'visio/pages/pages.xml';
	const pages = await zip.file(path)!.async('string');
	zip.file(path, pages.replace('N="DrawingScale" V="1"', `N="DrawingScale" V="${scale}"`));
	const { bytes } = await editVsdx(await zip.generateAsync({ type: 'uint8array' }), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 3, width: 2, height: 1 },
		{ type: 'create-rectangle', pageId: '0', shapeId: '2', x: 5, y: 4, width: 1, height: 1 },
	]);
	if (!lock) return bytes;
	const locked = await JSZip.loadAsync(bytes);
	const shapePath = 'visio/pages/page1.xml';
	locked.file(
		shapePath,
		(await locked.file(shapePath)!.async('string')).replace(
			'<Cell N="Width"',
			'<Cell N="LockWidth" V="1"/><Cell N="Width"',
		),
	);
	return locked.generateAsync({ type: 'uint8array' });
}
async function setup(options: { scale?: number; lock?: boolean; editor?: CancellableEditor } = {}) {
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		options.editor ?? edit,
		async (bytes, pageId, shapeIds) =>
			serializeVisioClipboard(await captureVisioClipboard(bytes, pageId, shapeIds)),
	);
	await controller.load(await fixture(options.scale, options.lock));
	controller.selectShape({ id: '1', name: 'Rectangle', pageId: '0' });
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.append(createSizePosition(document));
	const checked = vi.fn();
	const pane = new ViewerSizePosition(root, controller, checked);
	const dispose = pane.wire();
	const unsubscribe = controller.subscribe((state) => pane.render(state));
	pane.toggle(true);
	const input = (field: string) =>
		root.querySelector<HTMLInputElement>(`[data-size-field="${field}"]`)!;
	const set = (field: string, value: string) => {
		input(field).value = value;
		input(field).dispatchEvent(new Event('input'));
	};
	const enter = (field: string) =>
		input(field).dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
		);
	const done = () => vi.waitFor(() => expect(input('x').disabled).toBe(false));
	return {
		controller,
		root,
		pane,
		checked,
		input,
		set,
		enter,
		done,
		dispose: () => {
			dispose();
			unsubscribe();
			controller.destroy();
		},
	};
}
it.each([0.5, 1, 2])(
	'shows drawing quantities at scale %s, commits once, and restores exact source on undo',
	async (scale) => {
		const ui = await setup({ scale });
		expect(ui.input('x').value).toBe('2');
		expect(ui.input('width').value).toBe('2');
		const before = ui.controller.exportVsdx().bytes;
		const apply = vi.spyOn(ui.controller, 'applySelectionEdits');
		ui.set('x', '4');
		ui.enter('x');
		ui.input('x').dispatchEvent(new Event('blur'));
		await ui.done();
		expect(apply).toHaveBeenCalledOnce();
		expect(ui.controller.state.document!.pages[0]!.shapes[0]!.rotation!.pinX).toBeCloseTo(
			4 / scale,
			12,
		);
		expect(ui.input('x').value).toBe('4');
		await ui.controller.undo();
		expect(ui.controller.exportVsdx().bytes).toEqual(before);
		ui.dispose();
	},
);
it('commits dimensions with a fixed pin and degrees with the native positive angle sign', async () => {
	const ui = await setup();
	ui.set('width', '4');
	ui.enter('width');
	await ui.done();
	expect(ui.controller.state.document!.pages[0]!.shapes[0]).toMatchObject({
		width: 4,
		height: 1,
		rotation: { pinX: 2, pinY: 3 },
	});
	ui.set('angle', '-45');
	ui.input('angle').dispatchEvent(new Event('blur'));
	await ui.done();
	expect(ui.controller.state.document!.pages[0]!.shapes[0]!.rotation!.angle).toBeCloseTo(
		-Math.PI / 4,
		12,
	);
	expect(ui.input('angle').valueAsNumber).toBeCloseTo(-45, 12);
	ui.dispose();
});
it('retains other drafts through zoom and a committed edit, then resets on new selection or source', async () => {
	const ui = await setup();
	ui.set('height', '7');
	ui.controller.setZoom(2);
	expect(ui.input('height').value).toBe('7');
	ui.set('x', '4');
	ui.enter('x');
	await ui.done();
	expect(ui.input('height').value).toBe('7');
	ui.controller.selectShape({ id: '2', name: 'Other', pageId: '0' });
	expect(ui.input('height').value).toBe('1');
	ui.set('x', '8');
	await ui.controller.load(await fixture());
	expect(ui.input('x').value).toBe('');
	ui.dispose();
});
it('leaves source and history unchanged for untouched, equivalent, invalid and escaped values', async () => {
	const ui = await setup();
	const before = ui.controller.exportVsdx().bytes;
	const apply = vi.spyOn(ui.controller, 'applySelectionEdits');
	ui.enter('x');
	ui.set('x', '2.00');
	ui.enter('x');
	ui.set('width', '0');
	ui.enter('width');
	expect(ui.root.querySelector<HTMLElement>('[role=alert]')!.hidden).toBe(false);
	ui.set('width', '9');
	ui.input('width').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
	expect(ui.input('width').value).toBe('2');
	ui.input('width').dispatchEvent(new Event('blur'));
	expect(apply).not.toHaveBeenCalled();
	expect(ui.controller.exportVsdx().bytes).toEqual(before);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	ui.dispose();
});
it('shows authoritative protected-source rejection without changing source or history', async () => {
	const ui = await setup({ lock: true });
	const before = ui.controller.exportVsdx().bytes;
	ui.set('width', '4');
	ui.enter('width');
	await ui.done();
	expect(ui.root.querySelector<HTMLElement>('[role=alert]')!.hidden).toBe(false);
	expect(ui.controller.exportVsdx().bytes).toEqual(before);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	ui.dispose();
});
it('disables multiple, read-only and unsupported targets without using the primary selection alone', async () => {
	const ui = await setup();
	ui.controller.selectAll();
	expect(ui.input('x').disabled).toBe(true);
	expect(ui.input('x').value).toBe('');
	const model = structuredClone(ui.controller.state.document!);
	ui.controller.setDocument(model);
	ui.controller.selectShape({ id: '1', name: 'Read only', pageId: '0' });
	expect(ui.input('x').disabled).toBe(true);
	model.pages[0]!.shapes[0]!.kind = 'group';
	ui.controller.setDocument(model);
	ui.controller.selectShape({ id: '1', name: 'Group', pageId: '0' });
	expect(ui.input('width').value).toBe('');
	ui.dispose();
});
it('cancels stale deferred completion after selection changes and unwires disposal', async () => {
	let release!: () => void;
	const waiting = new Promise<void>((done) => {
		release = done;
	});
	const ui = await setup({
		editor: async (bytes, commands) => {
			await waiting;
			return edit(bytes, commands);
		},
	});
	const before = ui.controller.exportVsdx().bytes;
	ui.set('x', '6');
	ui.enter('x');
	ui.controller.clearSelection();
	release();
	await vi.waitFor(() => expect(ui.controller.state.edit.busy).toBe(false));
	expect(ui.controller.exportVsdx().bytes).toEqual(before);
	expect(ui.root.querySelector('[role=status]')!.textContent).toBe('');
	ui.dispose();
	ui.enter('x');
});
it('routes the shared ribbon pane action and close button through the mounted element', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host);
	const root = viewer.element.shadowRoot!;
	const pane = root.querySelector<HTMLElement>('.size-position')!;
	expect(pane.hidden).toBe(true);
	emitRibbonAction(root, { type: 'sizePosition' });
	expect(pane.hidden).toBe(false);
	expect(root.querySelector('[command="size-position"]')!.getAttribute('checked')).toBe('true');
	pane.querySelector<HTMLButtonElement>('button')!.click();
	expect(pane.hidden).toBe(true);
	expect(root.querySelector('[command="size-position"]')!.getAttribute('checked')).toBe('false');
	viewer.destroy();
});
