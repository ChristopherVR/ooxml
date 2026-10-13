import { afterEach, expect, it, vi } from 'vitest';
import { createVsdx, editVsdx, parseVsdx, type VisioEdit } from 'ooxml-core/visio';
import type { OfficeUiGallery } from '../ribbon/gallery';
import { ViewerController } from './controller';
import { ViewerCommands } from './viewer-commands';
import { createRibbon } from './ribbon';
import { createRulers } from './viewer-ruler';
import { registerViewerControls } from './office-ui';
import { renderPage } from './render-svg';
import { createContextMenus } from './viewer-context-menu';
import JSZip from 'jszip';

afterEach(() => document.body.replaceChildren());

async function setup(source = true) {
	registerViewerControls();
	const created = await createVsdx();
	const bytes = (
		await editVsdx(created, [
			{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 3, y: 4, width: 2, height: 1 },
		])
	).bytes;
	const edits: VisioEdit[][] = [];
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		async (input, commands) => {
			edits.push([...commands]);
			const result = await editVsdx(input, commands);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
	);
	if (source) await controller.load(bytes);
	else controller.setDocument(await parseVsdx(bytes));
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const viewport = document.createElement('div');
	viewport.className = 'viewport';
	root.append(createRibbon(document), viewport, ...createContextMenus(document));
	const feedback: string[] = [];
	const commands = new ViewerCommands({
		root,
		viewport,
		controller,
		rulers: createRulers(viewport),
		fit: () => {},
		togglePane: () => {},
		reveal: () => {},
		focusSearch: () => {},
		togglePanZoom: () => {},
		toggleSizePosition: () => {},
		announce: (message) => feedback.push(message),
	});
	const dispose = commands.wire();
	const paint = () => {
		const state = controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (page && state.document) viewport.replaceChildren(renderPage(state.document, page).svg);
		commands.render(state);
	};
	controller.subscribe(paint);
	paint();
	const done = () => vi.waitFor(() => expect(controller.state.edit.busy).toBe(false));
	const page = () => controller.state.document!.pages[controller.state.pageIndex]!;
	const run = (command: Parameters<ViewerCommands['run']>[0] & { type: 'page-setup' }) =>
		commands.run(command);
	return { root, viewport, controller, commands, edits, feedback, done, page, run, dispose, bytes };
}
const setupAction = (command: Record<string, unknown>) =>
	({ type: 'page-setup', command }) as Parameters<ViewerCommands['run']>[0] & {
		type: 'page-setup';
	};

it('edits saved jump settings through Layout and Routing as one undoable transaction', async () => {
	const ui = await setup();
	ui.run(setupAction({ op: 'dialog', tab: 'layout' }));
	const dialog = ui.root.querySelector('.page-setup-dialog')!;
	expect(dialog.querySelector('[data-setup-tab="layout"]')!.getAttribute('aria-selected')).toBe(
		'true',
	);
	for (const [field, value] of [
		['jump-code', '2'],
		['jump-style', '2'],
		['jump-direction-x', '2'],
		['jump-direction-y', '1'],
		['jump-size-x', '50'],
		['jump-size-y', '125'],
	])
		(
			dialog.querySelector(`[data-setup-field="${field}"]`) as HTMLElement & { value: string }
		).value = value!;
	dialog
		.querySelector('[command="page-setup-ok"]')!
		.dispatchEvent(new CustomEvent('office-command'));
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{
			type: 'set-page-layout',
			pageId: '0',
			lineJumpCode: 2,
			lineJumpStyle: 2,
			lineJumpDirectionX: 2,
			lineJumpDirectionY: 1,
			lineJumpFactorX: 0.5,
			lineJumpFactorY: 1.25,
		},
	]);
	expect(ui.page().layout).toMatchObject({ lineJumpDirectionX: 2, lineJumpDirectionY: 1 });
	await ui.controller.undo();
	expect(ui.page().layout?.lineJumpDirectionX).toBeUndefined();
	await ui.controller.redo();
	expect(ui.page().layout?.lineJumpDirectionX).toBe(2);
	ui.dispose();
});

it('enables Page Setup for a source drawing and explains the read-only case', async () => {
	const readOnly = await setup(false);
	const menu = readOnly.root.querySelector<HTMLElement & { disabled: boolean }>(
		'[data-menu="orientation"]',
	)!;
	expect(menu.disabled).toBe(true);
	expect(menu.title).toMatch(/Open a \.vsdx file/);
	readOnly.run(setupAction({ op: 'orientation', value: 'landscape' }));
	expect(readOnly.edits).toHaveLength(0);
	readOnly.dispose();
	const ui = await setup();
	expect(
		ui.root.querySelector<HTMLElement & { disabled: boolean }>('[data-menu="orientation"]')!
			.disabled,
	).toBe(false);
	expect(ui.root.querySelector('[command="orientation-portrait"]')!.getAttribute('checked')).toBe(
		'true',
	);
	expect(ui.root.querySelector('[command="size-letter"]')!.getAttribute('checked')).toBe('true');
	const group = ui.root.querySelector('office-ui-ribbon-group[launcher="page-setup-dialog"]')!;
	expect(group.hasAttribute('launcher-disabled')).toBe(false);
	ui.dispose();
});

it('changes orientation, size and Auto Size with undo', async () => {
	const ui = await setup();
	ui.run(setupAction({ op: 'orientation', value: 'landscape' }));
	await ui.done();
	expect([ui.page().width, ui.page().height]).toEqual([11, 8.5]);
	expect(ui.page().pageSetup?.printPageOrientation).toBe(2);
	expect(ui.root.querySelector('[command="orientation-landscape"]')!.getAttribute('checked')).toBe(
		'true',
	);
	ui.run(setupAction({ op: 'size', id: 'a4' }));
	await ui.done();
	expect(ui.page().width).toBeCloseTo(297 / 25.4);
	ui.run(setupAction({ op: 'auto-size' }));
	await ui.done();
	expect(ui.page().drawingResizeType).toBe(1);
	expect(ui.root.querySelector('[command="auto-size"]')!.getAttribute('pressed')).toBe('true');
	await ui.controller.undo();
	await ui.controller.undo();
	await ui.controller.undo();
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	ui.dispose();
});

it('fits the page to the drawing in one undoable step', async () => {
	const ui = await setup();
	ui.run(setupAction({ op: 'fit' }));
	await ui.done();
	expect(ui.edits.at(-1)!.map((edit) => edit.type)).toEqual(['set-page-size', 'move-shape']);
	expect([ui.page().width, ui.page().height]).toEqual([2, 1]);
	await ui.controller.undo();
	expect(ui.page().width).toBe(8.5);
	ui.dispose();
});

it('applies, recolours and removes backgrounds and borders', async () => {
	const ui = await setup();
	const backgrounds = ui.root.querySelector<OfficeUiGallery>(
		'office-ui-gallery[command="backgrounds"]',
	)!;
	expect(backgrounds.hasAttribute('disabled')).toBe(false);
	const color = ui.root.querySelector<HTMLElement & { disabled: boolean }>(
		'[data-menu="background-color"]',
	)!;
	expect(color.disabled).toBe(true);
	backgrounds.dispatchEvent(
		new CustomEvent('office-gallery-pick', {
			detail: { gallery: 'backgrounds', itemId: 'band' },
			bubbles: true,
		}),
	);
	await ui.done();
	expect(ui.controller.state.document!.pages.map((page) => page.name)).toEqual([
		'Page-1',
		'VBackground-1',
	]);
	expect(backgrounds.state!.sections[0]!.items.find((item) => item.id === 'band')!.applied).toBe(
		true,
	);
	expect(color.disabled).toBe(false);
	ui.run(setupAction({ op: 'background-color', color: '#FFF2CC' }));
	await ui.done();
	expect(ui.controller.state.document!.pages[1]!.shapes[0]!.style.fill.toUpperCase()).toBe(
		'#FFF2CC',
	);
	ui.run(setupAction({ op: 'border', style: 'simple' }));
	await ui.done();
	expect(
		ui.controller.state.document!.pages[1]!.shapes.find((shape) => shape.name === 'Title')!.text
			.plainText,
	).toBe('Page-1');
	ui.run(setupAction({ op: 'background', style: null }));
	await ui.done();
	ui.run(setupAction({ op: 'border', style: null }));
	await ui.done();
	expect(ui.controller.state.document!.pages).toHaveLength(1);
	ui.dispose();
});

it('draws page breaks from the print setup', async () => {
	const ui = await setup();
	await ui.controller.applyEdits([
		{ type: 'set-page-size', pageId: '0', width: 17, height: 11 },
		{ type: 'set-page-setup', pageId: '0', paperKind: 1, printOrientation: 'portrait' },
	]);
	ui.run(setupAction({ op: 'page-breaks' }));
	const lines = ui.viewport.querySelectorAll('[data-page-breaks] line');
	expect([...lines].map((line) => line.getAttribute('x1'))).toEqual(['8.5']);
	ui.run(setupAction({ op: 'page-breaks' }));
	expect(ui.viewport.querySelector('[data-page-breaks]')).toBeNull();
	ui.dispose();
});

it('draws line jumps where connectors cross and toggles them from Design > Connectors', async () => {
	const ui = await setup();
	const line = (shapeId: string, beginX: number, beginY: number, endX: number, endY: number) =>
		({
			type: 'create-line',
			pageId: '0',
			shapeId,
			beginX,
			beginY,
			endX,
			endY,
			route: 'straight',
		}) as VisioEdit;
	await ui.controller.applyEdits([line('2', 1, 2, 7, 2), line('3', 4, 1, 4, 6)]);
	const jumped = () => [...ui.viewport.querySelectorAll<SVGPathElement>('path[data-line-jumps]')];
	const item = ui.root.querySelector<HTMLElement>('[command="line-jumps"]')!;
	// Visio's default: the horizontal connector jumps over the vertical one with an arc.
	expect(jumped()).toHaveLength(1);
	expect(jumped()[0]!.getAttribute('d')).toMatch(/ A /);
	expect(item.getAttribute('checked')).toBe('true');
	expect(item.hasAttribute('disabled')).toBe(false);
	ui.run(setupAction({ op: 'line-jumps' }));
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([{ type: 'set-page-layout', pageId: '0', lineJumpCode: 0 }]);
	expect(jumped()).toHaveLength(0);
	expect(item.getAttribute('checked')).toBe('false');
	ui.run(setupAction({ op: 'line-jumps' }));
	await ui.done();
	expect(jumped()).toHaveLength(1);
	// Each toggle is one undoable step.
	await ui.controller.undo();
	expect(jumped()).toHaveLength(0);
	ui.dispose();
});

it('routes the page context menu to the existing dialog and cancels without edits', async () => {
	const ui = await setup();
	const menu = ui.root.querySelector<
		HTMLElement & { openAt(x: number, y: number): void; open: boolean }
	>('[data-context-menu="page"]')!;
	menu.openAt(5, 5);
	const button = menu.querySelector<HTMLElement & { disabled: boolean }>(
		'[command="ctx-page-setup"]',
	)!;
	expect(button.hasAttribute('data-unsupported')).toBe(false);
	expect(button.disabled).toBe(false);
	button.shadowRoot!.querySelector('button')!.click();
	const dialog = ui.root.querySelector<HTMLElement & { open: boolean }>('.page-setup-dialog')!;
	expect(menu.open).toBe(false);
	expect(dialog.open).toBe(true);
	expect(dialog.querySelector('[data-setup-tab="print"]')!.getAttribute('aria-selected')).toBe(
		'true',
	);
	dialog
		.querySelector<HTMLElement>('[command="page-setup-cancel"]')!
		.dispatchEvent(new CustomEvent('office-command', { bubbles: true }));
	expect(dialog.open).toBe(false);
	expect(ui.edits).toEqual([]);
	ui.dispose();
});

it('keeps context menu dialog routing inert without editable source or a document', async () => {
	const ui = await setup(false);
	const button = ui.root.querySelector('[command="ctx-page-setup"]')!;
	const dialog = ui.root.querySelector<HTMLElement & { open: boolean }>('.page-setup-dialog')!;
	button.shadowRoot!.querySelector('button')!.click();
	expect(dialog.open).toBe(false);
	ui.controller.setDocument(null);
	button.shadowRoot!.querySelector('button')!.click();
	expect(dialog.open).toBe(false);
	expect(ui.edits).toEqual([]);
	ui.dispose();
});

it('retains protected page settings when a context dialog edit is refused', async () => {
	const ui = await setup();
	const zip = await JSZip.loadAsync(ui.bytes);
	const path = 'visio/pages/pages.xml';
	const xml = await zip.file(path)!.async('string');
	const protectedXml = xml.replace(/(<Cell\b[^>]*N="PageWidth")/, '$1 F="GUARD(8.5)"');
	expect(protectedXml).not.toBe(xml);
	zip.file(path, protectedXml);
	await ui.controller.load(await zip.generateAsync({ type: 'uint8array' }));
	ui.root.querySelector('[command="ctx-page-setup"]')!.shadowRoot!.querySelector('button')!.click();
	const dialog = ui.root.querySelector<HTMLElement & { open: boolean }>('.page-setup-dialog')!;
	expect(dialog.open).toBe(true);
	dialog.querySelector<HTMLElement & { value: string }>('[data-setup-field="size-mode"]')!.value =
		'custom';
	dialog.querySelector<HTMLElement & { value: string }>('[data-setup-field="width"]')!.value = '17';
	dialog
		.querySelector<HTMLElement>('[command="page-setup-ok"]')!
		.dispatchEvent(new CustomEvent('office-command', { bubbles: true }));
	await vi.waitFor(() =>
		expect(dialog.querySelector('[data-setup-error]')!.textContent).toMatch(/protected/i),
	);
	expect(dialog.open).toBe(true);
	expect(ui.page().width).toBe(8.5);
	expect(ui.controller.state.edit.canUndo).toBe(false);
	ui.dispose();
});

it('commits the Page Setup dialog as one page transaction', async () => {
	const ui = await setup();
	ui.run(setupAction({ op: 'dialog', tab: 'scale' }));
	const dialog = ui.root.querySelector<HTMLElement & { open: boolean }>('.page-setup-dialog')!;
	expect(dialog.open).toBe(true);
	const field = (name: string) =>
		dialog.querySelector<HTMLElement & { value: string }>(`[data-setup-field="${name}"]`)!;
	expect(dialog.querySelector('[data-setup-panel="scale"]')!.hasAttribute('hidden')).toBe(false);
	field('scale-mode').value = 'custom';
	field('scale-mode').dispatchEvent(new Event('change'));
	field('scale-page').value = '1';
	field('scale-drawing').value = '1';
	field('scale-unit').value = 'FT';
	field('name').value = 'Floor plan';
	field('paper').value = '9';
	dialog
		.querySelector<HTMLElement>('[command="page-setup-ok"]')!
		.dispatchEvent(new CustomEvent('office-command', { detail: { command: 'page-setup-ok' } }));
	await vi.waitFor(() => expect(dialog.open).toBe(false));
	expect(ui.edits.at(-1)!.map((edit) => edit.type)).toEqual([
		'set-page-setup',
		'rename-page',
		'set-page-setup',
	]);
	expect(ui.page()).toMatchObject({ name: 'Floor plan', width: 8.5 });
	expect(ui.page().pageSetup).toMatchObject({ drawingScale: 12, paperKind: 9 });
	ui.dispose();
});
