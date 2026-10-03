import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountViewer } from './binding.js';
import { demoDocument } from './demo-document.js';

type Slider = HTMLElement & { value: number; disabled: boolean };
function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	const root = viewer.element.shadowRoot!;
	const button = (selector: string) => root.querySelector<HTMLButtonElement>(selector)!;
	/** Shared ribbon command host element and the real button inside it. */
	const command = (name: string) =>
		root.querySelector<HTMLElement & { disabled: boolean }>(`[command="${name}"]`)!;
	const press = (name: string) =>
		command(name).shadowRoot!.querySelector<HTMLButtonElement>('.main, button')!.click();
	const strip = () => root.querySelector('office-ui-tab-strip')!.shadowRoot!;
	const slider = () => root.querySelector<Slider>('office-ui-zoom-slider')!;
	const zoomButton = (label: string) =>
		slider().shadowRoot!.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`)!;
	return { host, viewer, root, button, command, press, strip, slider, zoomButton };
}
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

describe('shared Office-style viewer chrome', () => {
	it('uses one canvas and Visio page tabs, All pages and status instead of a page pane', () => {
		const { viewer, root, strip } = setup();
		expect(root.querySelectorAll('svg.paper')).toHaveLength(1);
		expect(root.querySelector('.page-rail, .page-link')).toBeNull();
		const all = root.querySelector('[data-menu="all-pages"]')!;
		const items = [...all.querySelectorAll('office-ui-menu-item')];
		expect(items.map((item) => item.getAttribute('label'))).toEqual([
			'Release workflow',
			'Architecture',
		]);
		expect(items[0]!.getAttribute('checked')).toBe('true');
		items[1]!.shadowRoot!.querySelector('button')!.click();
		expect(viewer.element.pageIndex).toBe(1);
		expect(items[1]!.getAttribute('checked')).toBe('true');
		expect(strip().querySelector('[aria-selected="true"]')!.textContent).toBe('Architecture');
		expect(root.querySelector('[data-page-status]')!.getAttribute('value')).toBe('Page 2 of 2');
		expect(root.querySelector('[data-page-name]')!.textContent).toBe('Architecture');
		strip().querySelectorAll<HTMLButtonElement>('[role="tab"]')[0]!.click();
		expect(viewer.element.pageIndex).toBe(0);
		strip().querySelector<HTMLButtonElement>('[aria-label="Next page"]')!.click();
		expect(viewer.element.pageIndex).toBe(1);
		// Insert Page is Visio's, shown disabled until core can add pages.
		const add = strip().querySelector<HTMLButtonElement>('.add')!;
		expect(add.hidden).toBe(false);
		expect(add.disabled).toBe(true);
		expect(add.title).toMatch(/not available yet\. Needs core page insertion/);
		const model = structuredClone(demoDocument);
		model.pages[0]!.name = '<img src=x onerror=alert(1)>';
		viewer.update({ document: model });
		expect(strip().querySelector('[role="tab"]')!.textContent).toBe(model.pages[0]!.name);
		expect(root.querySelector('img')).toBeNull();
		expect(strip().querySelector('img')).toBeNull();
		viewer.destroy();
	});
	it('opens the File backstage with real Info, Save and Close and disabled Visio pages', () => {
		const { viewer, root } = setup();
		const file = root.querySelector<HTMLButtonElement>('.file-tab')!;
		const backstage = root.querySelector<HTMLElement>('.backstage')!;
		expect(backstage.hidden).toBe(true);
		file.click();
		expect(backstage.hidden).toBe(false);
		expect(file.getAttribute('aria-expanded')).toBe('true');
		const info = root.querySelector<HTMLElement>('[data-backstage-page="info"]')!;
		expect(info.hidden).toBe(false);
		expect(info.querySelector('[data-info="pages"]')!.textContent).toBe('2');
		expect(info.querySelector('[data-info="state"]')!.textContent).toBe(
			'Model-only preview (read only)',
		);
		// The sample is model-only, so Save has nothing to save.
		expect(root.querySelector<HTMLButtonElement>('[data-backstage-item="save"]')!.disabled).toBe(
			true,
		);
		root.querySelector<HTMLButtonElement>('[data-backstage-item="new"]')!.click();
		const blank = root.querySelector<HTMLButtonElement>('[data-backstage-action="new-blank"]')!;
		expect(blank.disabled).toBe(true);
		expect(blank.title).toMatch(/not available yet\. Needs core blank drawing creation/);
		root.querySelector<HTMLButtonElement>('[data-backstage-item="export"]')!.click();
		expect(
			root.querySelector<HTMLButtonElement>(
				'[data-backstage-page="export"] [data-backstage-action="export-svg"]',
			)!.disabled,
		).toBe(false);
		backstage.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(backstage.hidden).toBe(true);
		expect(root.activeElement).toBe(file);
		file.click();
		root.querySelector<HTMLButtonElement>('[data-backstage-item="close"]')!.click();
		expect(viewer.element.document).toBeNull();
		expect(backstage.hidden).toBe(true);
		viewer.destroy();
	});
	it('provides real keyboard tabs and working pane toggles without changing zoom', () => {
		const { viewer, root, button, command, press } = setup();
		viewer.element.zoom = 1.7;
		button('[data-tab="home"]').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }),
		);
		expect(button('[data-tab="insert"]').getAttribute('aria-selected')).toBe('true');
		expect(root.activeElement).toBe(button('[data-tab="insert"]'));
		button('[data-tab="insert"]').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'End', bubbles: true }),
		);
		// Help is Visio's last tab; ArrowLeft steps back to View.
		expect(button('[data-tab="help"]').getAttribute('aria-selected')).toBe('true');
		button('[data-tab="help"]').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true }),
		);
		expect(button('[data-tab="view"]').getAttribute('aria-selected')).toBe('true');
		expect(root.activeElement).toBe(button('[data-tab="view"]'));
		expect(root.querySelector<HTMLElement>('#home-panel')!.hidden).toBe(true);
		expect(root.querySelector<HTMLElement>('#view-panel')!.hidden).toBe(false);
		// Visio's default: the Shapes window is open.
		const shapes = root.querySelector<HTMLElement>('.shapes-pane')!;
		expect(shapes.hidden).toBe(false);
		press('shapes');
		expect(shapes.hidden).toBe(true);
		expect(command('shapes').getAttribute('checked')).toBe('false');
		press('shapes');
		expect(shapes.hidden).toBe(false);
		press('inspector');
		expect(root.querySelector<HTMLElement>('.inspector-pane')!.hidden).toBe(true);
		expect(viewer.element.zoom).toBe(1.7);
		button('.notes-strip button').click();
		expect(root.querySelector<HTMLElement>('.inspector-pane')!.hidden).toBe(false);
		expect(command('inspector').getAttribute('checked')).toBe('true');
		expect(root.querySelector<HTMLDetailsElement>('.notes')!.open).toBe(true);
		expect(button('.notes-strip button').getAttribute('aria-expanded')).toBe('true');
		expect(root.activeElement).toBe(root.querySelector('.notes summary'));
		viewer.destroy();
	});
	it('drives zoom from the status slider and View commands and hides it with the toolbar', () => {
		const { viewer, root, press, slider, zoomButton } = setup();
		expect(slider().value).toBe(100);
		zoomButton('Zoom in').click();
		expect(viewer.element.zoom).toBeCloseTo(1.1);
		zoomButton('Zoom out').click();
		expect(viewer.element.zoom).toBe(1);
		const range = slider().shadowRoot!.querySelector('input')!;
		range.value = '250';
		range.dispatchEvent(new Event('input'));
		expect(viewer.element.zoom).toBe(2.5);
		press('zoom-100');
		expect(viewer.element.zoom).toBe(1);
		viewer.element.zoom = 2;
		expect(slider().value).toBe(200);
		viewer.update({ showToolbar: false });
		expect(root.querySelector<HTMLElement>('.toolbar')!.hidden).toBe(true);
		expect(slider().hidden).toBe(true);
		viewer.update({ showToolbar: true });
		expect(slider().hidden).toBe(false);
		viewer.destroy();
	});
	it('uses disabled states honestly and routes editing to the existing disclosure', () => {
		const { viewer, root, command, slider } = setup();
		expect(command('shape-data').disabled).toBe(true);
		expect(command('layer-properties').disabled).toBe(true);
		// The sample is a model-only document: no source bytes means no drawing or deletion.
		expect(command('rectangle').disabled).toBe(true);
		expect(command('undo').disabled).toBe(true);
		// Unsupported Visio commands are visible, disabled and say what is missing.
		expect(command('bold').disabled).toBe(true);
		expect(command('bold').getAttribute('title')).toMatch(/not available yet\. Needs core text/);
		root
			.querySelector<HTMLElement>('.viewport')!
			.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true, composed: true }));
		expect(root.querySelector<HTMLDetailsElement>('.edit-controls')!.open).toBe(true);
		expect(root.querySelector<HTMLTextAreaElement>('#edit-text')!.disabled).toBe(true);
		viewer.controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
		expect(command('shape-data').disabled).toBe(false);
		expect(root.querySelector<HTMLDetailsElement>('.shape-inspector')!.open).toBe(true);
		expect(root.querySelector('[data-shape-status]')!.getAttribute('value')).toMatch(/^Width: /);
		viewer.update({ document: null });
		expect(root.querySelectorAll('.page-link')).toHaveLength(0);
		expect(command('zoom-fit').disabled).toBe(true);
		expect(slider().disabled).toBe(true);
		expect(root.querySelector<HTMLButtonElement>('[data-tab="view"]')!.disabled).toBe(false);
		viewer.destroy();
	});
	it('aborts chrome, command and zoom listeners, then reconnects exactly one set', () => {
		const { host, viewer, root, button, press, strip, zoomButton } = setup();
		const view = button('[data-tab="view"]');
		const grid = () => root.querySelector<HTMLElement>('[data-check="grid"]')!.click();
		const next = () =>
			strip().querySelector<HTMLButtonElement>('[aria-label="Next page"]')!.click();
		viewer.element.remove();
		zoomButton('Zoom in').click();
		grid();
		view.click();
		expect(viewer.element.zoom).toBe(1);
		expect(root.querySelector<HTMLElement>('.viewport')!.dataset.grid).not.toBe('true');
		expect(view.getAttribute('aria-selected')).toBe('false');
		host.append(viewer.element);
		press('zoom-100');
		viewer.element.zoom = 1.25;
		grid();
		expect(root.querySelector<HTMLElement>('.viewport')!.dataset.grid).toBe('true');
		viewer.element.remove();
		host.append(viewer.element);
		grid();
		expect(root.querySelector<HTMLElement>('.viewport')!.dataset.grid).toBe('false');
		next();
		expect(viewer.element.pageIndex).toBe(1);
		const zoomIn = zoomButton('Zoom in');
		const inspector = root
			.querySelector('[command="inspector"]')!
			.shadowRoot!.querySelector('button')!;
		viewer.destroy();
		zoomIn.click();
		inspector.click();
		expect(root.children).toHaveLength(0);
	});
	it('fits to computed canvas padding and never enlarges the normal fit view', () => {
		const { viewer, root } = setup();
		const viewport = root.querySelector<HTMLElement>('.viewport')!;
		viewport.style.padding = '16px 4px';
		Object.defineProperty(viewport, 'clientWidth', { configurable: true, value: 1000 });
		Object.defineProperty(viewport, 'clientHeight', { configurable: true, value: 1000 });
		viewer.fit();
		expect(viewer.element.zoom).toBe(1);
		Object.defineProperty(viewport, 'clientWidth', { configurable: true, value: 416 });
		viewer.fit();
		expect(viewer.element.zoom).toBe(0.5);
		viewer.destroy();
	});
});
