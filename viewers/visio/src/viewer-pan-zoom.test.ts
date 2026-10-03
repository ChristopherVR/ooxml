import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountViewer } from './binding.js';
import { demoDocument } from './demo-document.js';

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: structuredClone(demoDocument) });
	const root = viewer.element.shadowRoot!;
	const press = (name: string) =>
		root
			.querySelector(`[command="${name}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('.main, button')!
			.click();
	return { viewer, root, press };
}

describe('Visio Shapes strip', () => {
	it('replaces the closed Shapes window with a strip that reopens it', () => {
		const { viewer, root } = setup();
		const pane = root.querySelector<HTMLElement>('.shapes-pane')!;
		const strip = root.querySelector<HTMLButtonElement>('.shapes-strip')!;
		expect(pane.hidden).toBe(false);
		expect(strip.hidden).toBe(true);
		root.querySelector<HTMLButtonElement>('.shapes-pane .pane-collapse')!.click();
		expect(pane.hidden).toBe(true);
		expect(strip.hidden).toBe(false);
		expect(strip.getAttribute('aria-label')).toBe('Open Shapes');
		strip.click();
		expect(pane.hidden).toBe(false);
		expect(strip.hidden).toBe(true);
		viewer.destroy();
	});
});

describe('Visio Pan & Zoom window', () => {
	it('opens from View > Task Panes with a page thumbnail and closes again', () => {
		const { viewer, root, press } = setup();
		const pane = root.querySelector<HTMLElement>('.pan-zoom')!;
		expect(pane.hidden).toBe(true);
		const item = root.querySelector('[command="pan-zoom"]')!;
		expect(item.hasAttribute('disabled')).toBe(false);
		press('pan-zoom');
		expect(pane.hidden).toBe(false);
		expect(item.getAttribute('checked')).toBe('true');
		const thumb = pane.querySelector('.pan-zoom-view svg')!;
		expect(thumb.getAttribute('aria-hidden')).toBe('true');
		expect(thumb.querySelector('.pan-zoom-frame')).not.toBeNull();
		expect(thumb.querySelector('[tabindex]')).toBeNull();
		pane.querySelector<HTMLButtonElement>('.pan-zoom-close')!.click();
		expect(pane.hidden).toBe(true);
		expect(item.getAttribute('checked')).toBe('false');
		viewer.destroy();
	});

	it('pans the drawing window with the arrow keys', () => {
		const { viewer, root, press } = setup();
		const viewport = root.querySelector<HTMLElement>('.viewport')!;
		const scrollBy = vi.fn();
		viewport.scrollBy = scrollBy as typeof viewport.scrollBy;
		press('pan-zoom');
		const view = root.querySelector<HTMLElement>('.pan-zoom-view')!;
		view.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		expect(scrollBy).toHaveBeenCalledWith({ left: 48, top: 0 });
		viewer.destroy();
	});
});
