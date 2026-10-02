import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountViewer } from './binding.js';
import { demoDocument } from './demo-document.js';

function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	const root = viewer.element.shadowRoot!;
	const button = (selector: string) => root.querySelector<HTMLButtonElement>(selector)!;
	return { host, viewer, root, button };
}
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

describe('shared Office-style viewer chrome', () => {
	it('uses one canvas, literal page cards, and synchronized keyboard navigation', () => {
		const { viewer, root, button } = setup();
		expect(root.querySelectorAll('svg')).toHaveLength(1);
		expect(root.querySelectorAll('.page-link')).toHaveLength(2);
		const first = button('[data-page-index="0"]'),
			second = button('[data-page-index="1"]');
		expect(first.getAttribute('aria-current')).toBe('page');
		second.click();
		expect(viewer.element.pageIndex).toBe(1);
		expect(root.querySelector('select')!.value).toBe('1');
		expect(second.getAttribute('aria-current')).toBe('page');
		expect(first.hasAttribute('aria-current')).toBe(false);
		expect(root.querySelector('[data-page-name]')!.textContent).toBe('Architecture');
		second.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
		expect(viewer.element.pageIndex).toBe(0);
		expect(root.activeElement).toBe(first);
		const model = structuredClone(demoDocument);
		model.pages[0]!.name = '<img src=x onerror=alert(1)>';
		viewer.update({ document: model });
		expect(root.querySelector('.page-name')!.textContent).toBe(model.pages[0]!.name);
		expect(root.querySelector('img')).toBeNull();
		expect(root.querySelectorAll('svg')).toHaveLength(1);
		viewer.destroy();
	});
	it('provides real keyboard tabs and working pane toggles without changing zoom', () => {
		const { viewer, root, button } = setup();
		viewer.element.zoom = 1.7;
		button('[data-tab="home"]').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }),
		);
		expect(button('[data-tab="view"]').getAttribute('aria-selected')).toBe('true');
		expect(root.activeElement).toBe(button('[data-tab="view"]'));
		expect(root.querySelector<HTMLElement>('#home-panel')!.hidden).toBe(true);
		expect(root.querySelector<HTMLElement>('#view-panel')!.hidden).toBe(false);
		button('[data-chrome="pages"]').click();
		expect(root.querySelector<HTMLElement>('.page-rail')!.hidden).toBe(true);
		expect(button('[data-chrome="pages"]').getAttribute('aria-pressed')).toBe('false');
		button('[data-chrome="inspector"]').click();
		expect(root.querySelector<HTMLElement>('.inspector-pane')!.hidden).toBe(true);
		expect(viewer.element.zoom).toBe(1.7);
		button('.notes-strip button').click();
		expect(root.querySelector<HTMLElement>('.inspector-pane')!.hidden).toBe(false);
		expect(root.querySelector<HTMLDetailsElement>('.notes')!.open).toBe(true);
		expect(button('.notes-strip button').getAttribute('aria-expanded')).toBe('true');
		expect(root.activeElement).toBe(root.querySelector('.notes summary'));
		viewer.destroy();
	});
	it('keeps status zoom commands working on either tab and hides them with the toolbar', () => {
		const { viewer, root, button } = setup();
		button('[data-tab="view"]').click();
		button('[data-action="in"]').click();
		expect(viewer.element.zoom).toBe(1.25);
		button('[data-action="out"]').click();
		expect(viewer.element.zoom).toBe(1);
		viewer.element.zoom = 2;
		button('[data-action="actual"]').click();
		expect(viewer.element.zoom).toBe(1);
		viewer.update({ showToolbar: false });
		expect(root.querySelector<HTMLElement>('.toolbar')!.hidden).toBe(true);
		expect(root.querySelector<HTMLElement>('.zoom-controls')!.hidden).toBe(true);
		viewer.update({ showToolbar: true });
		expect(root.querySelector<HTMLElement>('.zoom-controls')!.hidden).toBe(false);
		viewer.destroy();
	});
	it('uses disabled states honestly and routes editing to the existing disclosure', () => {
		const { viewer, root, button } = setup();
		expect(button('[data-chrome="selection"]').disabled).toBe(true);
		expect(button('[data-chrome="layers"]').disabled).toBe(true);
		button('[data-chrome="edit"]').click();
		expect(root.querySelector<HTMLDetailsElement>('.edit-controls')!.open).toBe(true);
		expect(root.querySelector<HTMLTextAreaElement>('#edit-text')!.disabled).toBe(true);
		viewer.controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
		expect(button('[data-chrome="selection"]').disabled).toBe(false);
		expect(root.querySelector<HTMLDetailsElement>('.shape-inspector')!.open).toBe(true);
		viewer.update({ document: null });
		expect(root.querySelectorAll('.page-link')).toHaveLength(0);
		expect(button('[data-chrome="edit"]').disabled).toBe(true);
		expect(button('[data-action="in"]').disabled).toBe(true);
		expect(button('[data-tab="view"]').disabled).toBe(false);
		viewer.destroy();
	});
	it('aborts chrome and moved zoom listeners, then reconnects exactly one set', () => {
		const { host, viewer, root, button } = setup();
		const zoom = button('[data-action="in"]'),
			page = button('[data-chrome="next-page"]');
		const view = button('[data-tab="view"]');
		viewer.element.remove();
		zoom.click();
		page.click();
		view.click();
		expect(viewer.element.zoom).toBe(1);
		expect(viewer.element.pageIndex).toBe(0);
		expect(view.getAttribute('aria-selected')).toBe('false');
		host.append(viewer.element);
		zoom.click();
		expect(viewer.element.zoom).toBe(1.25);
		viewer.element.remove();
		host.append(viewer.element);
		zoom.click();
		expect(viewer.element.zoom).toBe(1.5625);
		page.click();
		expect(viewer.element.pageIndex).toBe(1);
		const toggle = button('[data-chrome="inspector"]');
		viewer.destroy();
		zoom.click();
		page.click();
		toggle.click();
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
