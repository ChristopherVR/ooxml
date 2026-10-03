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
		root.querySelector<HTMLElement & { disabled: boolean }>(`office-ui-button[command="${name}"]`)!;
	const press = (name: string) => command(name).shadowRoot!.querySelector('button')!.click();
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
	it('uses one canvas, literal page cards, page tabs and synchronized keyboard navigation', () => {
		const { viewer, root, button, strip } = setup();
		expect(root.querySelectorAll('svg.paper')).toHaveLength(1);
		expect(root.querySelectorAll('.page-link')).toHaveLength(2);
		const first = button('[data-page-index="0"]'),
			second = button('[data-page-index="1"]');
		expect(first.getAttribute('aria-current')).toBe('page');
		second.click();
		expect(viewer.element.pageIndex).toBe(1);
		expect(strip().querySelector('[aria-selected="true"]')!.textContent).toBe('Architecture');
		expect(root.querySelector('[data-page-status]')!.getAttribute('value')).toBe('Page 2 of 2');
		expect(second.getAttribute('aria-current')).toBe('page');
		expect(first.hasAttribute('aria-current')).toBe(false);
		expect(root.querySelector('[data-page-name]')!.textContent).toBe('Architecture');
		second.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
		expect(viewer.element.pageIndex).toBe(0);
		expect(root.activeElement).toBe(first);
		strip().querySelectorAll<HTMLButtonElement>('[role="tab"]')[1]!.click();
		expect(viewer.element.pageIndex).toBe(1);
		strip().querySelector<HTMLButtonElement>('[aria-label="Previous page"]')!.click();
		expect(viewer.element.pageIndex).toBe(0);
		const model = structuredClone(demoDocument);
		model.pages[0]!.name = '<img src=x onerror=alert(1)>';
		viewer.update({ document: model });
		expect(root.querySelector('.page-name')!.textContent).toBe(model.pages[0]!.name);
		expect(strip().querySelector('[role="tab"]')!.textContent).toBe(model.pages[0]!.name);
		expect(root.querySelector('img')).toBeNull();
		expect(strip().querySelector('img')).toBeNull();
		viewer.destroy();
	});
	it('provides real keyboard tabs and working pane toggles without changing zoom', () => {
		const { viewer, root, button, command, press } = setup();
		viewer.element.zoom = 1.7;
		button('[data-tab="home"]').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }),
		);
		expect(button('[data-tab="view"]').getAttribute('aria-selected')).toBe('true');
		expect(root.activeElement).toBe(button('[data-tab="view"]'));
		expect(root.querySelector<HTMLElement>('#home-panel')!.hidden).toBe(true);
		expect(root.querySelector<HTMLElement>('#view-panel')!.hidden).toBe(false);
		press('pages');
		expect(root.querySelector<HTMLElement>('.page-rail')!.hidden).toBe(true);
		expect(command('pages').getAttribute('pressed')).toBe('false');
		press('inspector');
		expect(root.querySelector<HTMLElement>('.inspector-pane')!.hidden).toBe(true);
		expect(viewer.element.zoom).toBe(1.7);
		button('.notes-strip button').click();
		expect(root.querySelector<HTMLElement>('.inspector-pane')!.hidden).toBe(false);
		expect(command('inspector').getAttribute('pressed')).toBe('true');
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
		press('actual-size');
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
		const { viewer, root, command, press, slider } = setup();
		expect(command('selection').disabled).toBe(true);
		expect(command('layers').disabled).toBe(true);
		// The sample is a model-only document: no source bytes means no drawing or deletion.
		expect(command('rectangle').disabled).toBe(true);
		expect(command('undo').disabled).toBe(true);
		press('edit');
		expect(root.querySelector<HTMLDetailsElement>('.edit-controls')!.open).toBe(true);
		expect(root.querySelector<HTMLTextAreaElement>('#edit-text')!.disabled).toBe(true);
		viewer.controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
		expect(command('selection').disabled).toBe(false);
		expect(command('delete').disabled).toBe(true);
		expect(root.querySelector<HTMLDetailsElement>('.shape-inspector')!.open).toBe(true);
		expect(root.querySelector('[data-shape-status]')!.getAttribute('value')).toMatch(/^Width: /);
		viewer.update({ document: null });
		expect(root.querySelectorAll('.page-link')).toHaveLength(0);
		expect(command('edit').disabled).toBe(true);
		expect(command('zoom-fit').disabled).toBe(true);
		expect(slider().disabled).toBe(true);
		expect(root.querySelector<HTMLButtonElement>('[data-tab="view"]')!.disabled).toBe(false);
		viewer.destroy();
	});
	it('aborts chrome, command and zoom listeners, then reconnects exactly one set', () => {
		const { host, viewer, root, button, press, strip, zoomButton } = setup();
		const view = button('[data-tab="view"]');
		const next = () =>
			strip().querySelector<HTMLButtonElement>('[aria-label="Next page"]')!.click();
		viewer.element.remove();
		zoomButton('Zoom in').click();
		press('grid');
		view.click();
		expect(viewer.element.zoom).toBe(1);
		expect(root.querySelector<HTMLElement>('.viewport')!.dataset.grid).not.toBe('true');
		expect(view.getAttribute('aria-selected')).toBe('false');
		host.append(viewer.element);
		press('actual-size');
		viewer.element.zoom = 1.25;
		press('grid');
		expect(root.querySelector<HTMLElement>('.viewport')!.dataset.grid).toBe('true');
		viewer.element.remove();
		host.append(viewer.element);
		press('grid');
		expect(root.querySelector<HTMLElement>('.viewport')!.dataset.grid).toBe('false');
		next();
		expect(viewer.element.pageIndex).toBe(1);
		const zoomIn = zoomButton('Zoom in');
		const inspector = root
			.querySelector('office-ui-button[command="inspector"]')!
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
