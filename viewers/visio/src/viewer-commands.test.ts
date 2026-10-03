import { afterEach, describe, expect, it } from 'vitest';
import type { VisioEdit } from 'ooxml-core/visio';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';
import { registerViewerControls } from './office-ui.js';
import { ViewerCommands } from './viewer-commands.js';
import { nextShapeId } from './viewer-draw-tool.js';
import { createRibbon } from './ribbon.js';
import { createRulers } from './viewer-ruler.js';
import type { CancellableEditor } from './worker-editor.js';

afterEach(() => document.body.replaceChildren());

async function setup(source = true) {
	registerViewerControls();
	const edits: VisioEdit[][] = [];
	const editor: CancellableEditor = async (_bytes, commands) => {
		edits.push([...commands]);
		return {
			bytes: new Uint8Array([edits.length + 1]),
			document: structuredClone(demoDocument),
			changedParts: ['visio/pages/page1.xml'],
			diagnostics: [],
		};
	};
	const controller = new ViewerController(
		async () => structuredClone(demoDocument),
		() => {},
		editor,
	);
	if (source) await controller.load(new Uint8Array([1]));
	else controller.setDocument(structuredClone(demoDocument));
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const viewport = document.createElement('div');
	viewport.className = 'viewport';
	viewport.tabIndex = 0;
	viewport.append(document.createElement('textarea'));
	root.append(createRibbon(document), viewport);
	const rulers = createRulers(viewport);
	const calls: string[] = [];
	const commands = new ViewerCommands({
		root,
		viewport,
		rulers,
		controller,
		fit: (mode) => calls.push(`fit:${mode}`),
		togglePane: (pane) => calls.push(`pane:${pane}`),
		reveal: (panel, focusText) => calls.push(`reveal:${panel}:${focusText}`),
		focusSearch: () => calls.push('search'),
		announce: (message) => calls.push(message),
	});
	const dispose = commands.wire();
	controller.subscribe((state) => commands.render(state));
	/** Any ribbon command, split menu or menu item by stable id. */
	const command = (name: string) =>
		root.querySelector<HTMLElement & { disabled: boolean }>(`[command="${name}"]`)!;
	const press = (name: string) =>
		command(name).shadowRoot!.querySelector<HTMLButtonElement>('.main, button')!.click();
	const check = (name: string) =>
		root.querySelector<HTMLElement & { checked: boolean }>(`[data-check="${name}"]`)!;
	const key = (init: KeyboardEventInit, target: Element = viewport) =>
		target.dispatchEvent(
			new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }),
		);
	const settle = () => new Promise((resolve) => setTimeout(resolve, 0));
	return {
		controller,
		root,
		viewport,
		commands,
		calls,
		edits,
		command,
		press,
		check,
		key,
		settle,
		dispose,
	};
}

describe('Visio ribbon commands', () => {
	it('turns shared button activation into typed ribbon actions', async () => {
		const { root, press } = await setup();
		const actions: unknown[] = [];
		const raw: unknown[] = [];
		root.addEventListener('ribbon-action', (event) => actions.push((event as CustomEvent).detail));
		root.addEventListener('office-command', (event) => raw.push((event as CustomEvent).detail));
		press('rectangle');
		press('page-width');
		press('inspector');
		expect(actions).toEqual([
			{ type: 'tool', tool: 'rectangle' },
			{ type: 'zoom', mode: 'width' },
			{ type: 'pane', pane: 'inspector' },
		]);
		expect(raw).toEqual([]);
	});

	it('routes shared office-command buttons to pane, zoom and tool commands', async () => {
		const { calls, command, press, check, viewport, controller } = await setup();
		press('shapes');
		press('layer-properties');
		press('zoom-fit');
		press('page-width');
		expect(calls).toEqual(['pane:shapes', 'reveal:layers:false', 'fit:page', 'fit:width']);
		press('rectangle');
		expect(command('rectangle').hasAttribute('data-active')).toBe(true);
		expect(command('rectangle-item').getAttribute('checked')).toBe('true');
		expect(command('pointer').getAttribute('pressed')).toBe('false');
		expect(viewport.dataset.tool).toBe('rectangle');
		check('grid').click();
		expect(viewport.dataset.grid).toBe('true');
		expect(check('grid').checked).toBe(true);
		check('ruler').click();
		expect(viewport.closest('.canvas-area')!.getAttribute('data-ruler')).toBe('true');
		press('zoom-150');
		expect(controller.state.zoom).toBe(1.5);
	});

	it('deletes the selection with Delete, then undoes and redoes with Visio shortcuts', async () => {
		const { controller, command, edits, key, settle, calls } = await setup();
		expect(command('undo').disabled).toBe(true);
		controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
		key({ key: 'Delete' });
		await settle();
		expect(edits).toEqual([[{ type: 'delete-shape', pageId: '1', shapeId: 's1' }]]);
		expect(calls).toContain('Deleted Start.');
		expect(command('undo').disabled).toBe(false);
		key({ key: 'z', ctrlKey: true });
		await settle();
		expect(controller.state.edit.canRedo).toBe(true);
		key({ key: 'y', ctrlKey: true });
		await settle();
		expect(controller.state.edit).toMatchObject({ canUndo: true, canRedo: false });
	});

	it('maps Visio tool, page, find, fit and text shortcuts', async () => {
		const { controller, command, key, calls } = await setup();
		key({ key: '8', ctrlKey: true });
		expect(command('rectangle').hasAttribute('data-active')).toBe(true);
		key({ key: 'Escape' });
		expect(command('pointer').getAttribute('pressed')).toBe('true');
		key({ key: 'PageDown', ctrlKey: true });
		expect(controller.state.pageIndex).toBe(1);
		key({ key: 'PageUp', ctrlKey: true });
		expect(controller.state.pageIndex).toBe(0);
		key({ key: 'F2' });
		key({ key: 'f', ctrlKey: true });
		key({ key: 'W', ctrlKey: true, shiftKey: true });
		expect(calls).toEqual(['reveal:edit:true', 'search', 'fit:page']);
	});

	it('leaves undo and deletion keys to text fields', async () => {
		const { controller, root, key, edits, settle } = await setup();
		controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
		const text = root.querySelector('textarea')!;
		expect(key({ key: 'Delete' }, text)).toBe(true);
		expect(key({ key: 'z', ctrlKey: true }, text)).toBe(true);
		await settle();
		expect(edits).toEqual([]);
	});

	it('keeps drawing and deletion unavailable for model-only documents', async () => {
		const { controller, command, key, edits, settle } = await setup(false);
		expect(command('rectangle').disabled).toBe(true);
		expect(command('rectangle').title).toMatch(/model-only/i);
		key({ key: '8', ctrlKey: true });
		expect(command('pointer').getAttribute('pressed')).toBe('true');
		controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
		key({ key: 'Delete' });
		key({ key: 'z', ctrlKey: true });
		await settle();
		expect(edits).toEqual([]);
	});

	it('stops listening after disposal', async () => {
		const { calls, press, dispose } = await setup();
		dispose();
		press('zoom-fit');
		expect(calls).toEqual([]);
	});
});

describe('rectangle tool shape IDs', () => {
	it('allocates the next numeric ID across nested shapes', () => {
		const page = structuredClone(demoDocument.pages[0]!);
		page.shapes[0]!.id = '7';
		page.shapes[0]!.children = [{ ...structuredClone(page.shapes[1]!), id: '41', children: [] }];
		expect(nextShapeId(page)).toBe('42');
		expect(nextShapeId({ ...page, shapes: [] })).toBe('1');
	});
});

describe('Visio context menus', () => {
	it('opens the shape menu on right-click, selects the shape and runs Edit Text', async () => {
		const { mountViewer } = await import('./binding.js');
		const host = document.createElement('div');
		document.body.append(host);
		const viewer = mountViewer(host, { document: structuredClone(demoDocument) });
		const root = viewer.element.shadowRoot!;
		const shape = root.querySelector<SVGGElement>('svg.paper [data-shape-id="s1"]')!;
		shape.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 30, clientY: 40 }));
		const menu = root.querySelector<HTMLElement & { open: boolean }>(
			'[data-context-menu="shape"]',
		)!;
		expect(menu.open).toBe(true);
		expect(viewer.controller.state.selectedShape?.id).toBe('s1');
		const cut = menu.querySelector<HTMLElement & { disabled: boolean }>('[command="ctx-cut"]')!;
		expect(cut.disabled).toBe(true);
		expect(cut.getAttribute('title')).toMatch(/not available yet\. Needs core shape copy/);
		menu.querySelector('[command="ctx-edit-text"]')!.shadowRoot!.querySelector('button')!.click();
		expect(menu.open).toBe(false);
		expect(root.querySelector<HTMLDetailsElement>('.edit-controls')!.open).toBe(true);
		root
			.querySelector('.viewport')!
			.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 5, clientY: 5 }));
		const page = root.querySelector<HTMLElement & { open: boolean }>('[data-context-menu="page"]')!;
		expect(page.open).toBe(true);
		page.querySelector('[command="ctx-grid"]')!.shadowRoot!.querySelector('button')!.click();
		expect(root.querySelector<HTMLElement>('.viewport')!.dataset.grid).toBe('true');
		viewer.destroy();
	});
});
