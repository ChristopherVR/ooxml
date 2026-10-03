import { afterEach, describe, expect, it } from 'vitest';
import { mountViewer, type MountedViewer } from './binding.js';
import { demoDocument } from './demo-document.js';

const mounted: MountedViewer[] = [];
afterEach(() => {
	for (const viewer of mounted.splice(0)) viewer.destroy();
	document.body.replaceChildren();
});
function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: structuredClone(demoDocument) });
	mounted.push(viewer);
	const root = viewer.element.shadowRoot!;
	const button = (selector: string) => root.querySelector<HTMLButtonElement>(selector)!;
	return { viewer, root, button };
}
describe('independent shared-chrome regression review', () => {
	it('inserts hostile page names as text, including updated document replacements', () => {
		const { viewer, root } = setup();
		const doc = structuredClone(demoDocument);
		doc.pages[0]!.name = '<img src=x onerror=alert(1)> & "Page"';
		viewer.update({ document: doc });
		expect(root.querySelector('.page-name')!.textContent).toBe(doc.pages[0]!.name);
		expect(root.querySelector('[data-page-name]')!.textContent).toBe(doc.pages[0]!.name);
		expect(root.querySelector('.page-list img')).toBeNull();
		expect(root.querySelector('.page-list [onerror]')).toBeNull();
	});
	it('keeps page tabs, page rail, stepper, status and inspector on the same real page', () => {
		const { viewer, root, button } = setup();
		const strip = root.querySelector('office-ui-tab-strip')!.shadowRoot!;
		const step = (label: string) =>
			strip.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`)!;
		step('Next page').click();
		expect(viewer.controller.state.pageIndex).toBe(1);
		expect(strip.querySelector('[aria-selected="true"]')!.textContent).toBe(
			demoDocument.pages[1]!.name,
		);
		expect(root.querySelector('[data-page-status]')!.getAttribute('value')).toBe('Page 2 of 2');
		expect(root.querySelector('[aria-current="page"]')!.getAttribute('data-page-index')).toBe('1');
		expect(root.querySelector('[data-page-name]')!.textContent).toBe(demoDocument.pages[1]!.name);
		expect(step('Next page').disabled).toBe(true);
		button('[data-page-index="0"]').click();
		expect(viewer.controller.state.pageIndex).toBe(0);
		expect(step('Previous page').disabled).toBe(true);
	});
	it('moves tab focus and panel visibility together for keyboard navigation', () => {
		const { root, button } = setup();
		button('[data-tab="home"]').focus();
		button('[data-tab="home"]').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'End', bubbles: true }),
		);
		expect(root.activeElement).toBe(button('[data-tab="view"]'));
		expect(button('[data-tab="view"]').getAttribute('aria-selected')).toBe('true');
		expect(root.querySelector<HTMLElement>('#home-panel')!.hidden).toBe(true);
		expect(root.querySelector<HTMLElement>('#view-panel')!.hidden).toBe(false);
		button('[data-tab="view"]').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Home', bubbles: true }),
		);
		expect(root.activeElement).toBe(button('[data-tab="home"]'));
	});
	it('preserves zoom and document while toggling panes and toolbar', () => {
		const { viewer, root, button } = setup();
		viewer.update({ zoom: 1.75 });
		const original = viewer.controller.state.document;
		button('[data-chrome="inspector"]').click();
		expect(root.querySelector<HTMLElement>('.inspector-pane')!.hidden).toBe(true);
		button('[data-chrome="inspector"]').click();
		viewer.update({ showToolbar: false });
		expect(root.querySelector<HTMLElement>('.toolbar')!.hidden).toBe(true);
		expect(root.querySelector<HTMLElement>('.zoom-controls')!.hidden).toBe(true);
		viewer.update({ showToolbar: true });
		expect(root.querySelector<HTMLElement>('.zoom-controls')!.hidden).toBe(false);
		expect(viewer.controller.state.document).toBe(original);
		expect(viewer.controller.state.zoom).toBe(1.75);
	});
});
