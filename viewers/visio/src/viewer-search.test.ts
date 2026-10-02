import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountViewer } from './binding.js';
import { demoDocument } from './demo-document.js';
import { createVsdxFixture } from '../tests/fixture.mjs';

afterEach(() => document.body.replaceChildren());
function mounted(connected = true) {
	const host = document.createElement('div');
	if (connected) document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	const root = viewer.element.shadowRoot!;
	const input = root.querySelector<HTMLInputElement>('input[type="search"]')!;
	const status = root.querySelector<HTMLElement>('#search-status')!;
	const next = root.querySelector<HTMLButtonElement>('[data-action="search-next"]')!;
	const previous = root.querySelector<HTMLButtonElement>('[data-action="search-previous"]')!;
	const query = (value: string) => {
		input.value = value;
		input.dispatchEvent(new Event('input', { bubbles: true }));
	};
	return { host, viewer, root, input, status, next, previous, query };
}
describe('accessible shared document text search UI', () => {
	it('searches imported canonical text through the local parser and shared renderer', async () => {
		const { viewer, query, next, root } = mounted();
		await viewer.load(await createVsdxFixture('<script>Searchable text</script>'));
		query('searchable');
		expect(viewer.controller.state.search.results).toHaveLength(1);
		next.click();
		const result = viewer.controller.state.search.results[0]!;
		expect(viewer.controller.state.selectedShape?.pageId).toBe(result.pageId);
		expect(root.querySelector('[data-selected="true"]')?.textContent).toContain('Searchable text');
		expect(root.querySelector('script')).toBeNull();
		viewer.destroy();
	});
	it('labels bounded input, navigates matching shapes, and retains keyboard focus', () => {
		const { viewer, input, status, next, previous, query, root } = mounted();
		expect(input.getAttribute('aria-label')).toBe('Search diagram text');
		expect(input.maxLength).toBe(256);
		expect(input.getAttribute('aria-describedby')).toBe(status.id);
		expect(status.getAttribute('aria-live')).toBe('polite');
		expect(next.disabled).toBe(true);
		expect(previous.disabled).toBe(true);
		input.focus();
		query('framework');
		expect(status.textContent).toBe('2 matching shapes');
		expect(viewer.element.pageIndex).toBe(0);
		expect(next.disabled).toBe(false);
		input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(viewer.element.pageIndex).toBe(1);
		expect(viewer.controller.state.selectedShape?.id).toBe('a1');
		expect(status.textContent).toBe('1 of 2 matching shapes');
		expect(root.activeElement).toBe(input);
		expect(root.querySelector('[data-selected="true"]')?.getAttribute('data-shape-id')).toBe('a1');
		input.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true, bubbles: true }),
		);
		expect(viewer.controller.state.selectedShape?.id).toBe('a4');
		expect(status.textContent).toBe('2 of 2 matching shapes');
		expect(status.title).toContain('Architecture');
		next.click();
		expect(viewer.controller.state.selectedShape?.id).toBe('a1');
		previous.click();
		expect(viewer.controller.state.selectedShape?.id).toBe('a4');
		input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(input.value).toBe('');
		expect(next.disabled).toBe(true);
		expect(status.textContent).toBe('Find text across pages');
		viewer.destroy();
	});
	it('does not redraw or rescan SVG selection while typing a query', () => {
		const { viewer, root, query } = mounted();
		viewer.controller.selectShape({ id: 's1', name: 'Start', pageId: '1' });
		query('idea');
		const viewport = root.querySelector<HTMLDivElement>('.viewport')!;
		const svg = viewport.querySelector('svg');
		const scan = vi.spyOn(viewport, 'querySelectorAll');
		query('fram');
		query('frame');
		query('framework');
		expect(scan).not.toHaveBeenCalled();
		expect(viewport.querySelector('svg')).toBe(svg);
		expect(viewport.querySelector('[data-selected="true"]')?.getAttribute('data-shape-id')).toBe(
			's1',
		);
		viewer.destroy();
	});
	it('does not navigate during IME composition or interpret input as viewport shortcuts', () => {
		const { viewer, input, query } = mounted();
		query('framework');
		input.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', isComposing: true, bubbles: true }),
		);
		input.dispatchEvent(new KeyboardEvent('keydown', { key: '+', bubbles: true }));
		expect(viewer.controller.state.search.activeIndex).toBe(-1);
		expect(viewer.element.zoom).toBe(1);
		viewer.destroy();
	});
	it('scrolls a selected result into view without looking up unsafe shape identities', () => {
		const { viewer, root, query, next } = mounted();
		query('idea');
		const shape = root.querySelector<SVGGElement>('[data-shape-id="s1"]')!;
		const scroll = vi.fn();
		shape.scrollIntoView = scroll;
		next.click();
		expect(scroll).toHaveBeenCalledWith({ block: 'nearest', inline: 'nearest' });
		viewer.controller.setZoom(2);
		expect(scroll).toHaveBeenCalledOnce();
		query('start');
		next.click();
		expect(scroll).toHaveBeenCalledTimes(2);
		viewer.destroy();
	});
	it('preserves disconnected search state, aborts inputs, and reconnects one listener set', () => {
		const { host, viewer, query, next, input } = mounted();
		query('framework');
		const search = viewer.controller.state.search;
		viewer.element.remove();
		next.click();
		query('disconnected');
		expect(viewer.controller.state.search).toBe(search);
		host.append(viewer.element);
		expect(input.value).toBe('framework');
		next.click();
		expect(viewer.controller.state.search.activeIndex).toBe(0);
		viewer.element.remove();
		host.append(viewer.element);
		next.click();
		expect(viewer.controller.state.search.activeIndex).toBe(1);
		viewer.destroy();
	});
	it('synchronizes imperative queries and resets UI on replacement or clearing the document', () => {
		const { viewer, input, status, next } = mounted();
		viewer.controller.setSearchQuery('framework');
		expect(input.value).toBe('framework');
		viewer.update({ document: { ...demoDocument } });
		expect(input.value).toBe('');
		expect(next.disabled).toBe(true);
		viewer.update({ document: null });
		expect(input.disabled).toBe(true);
		expect(status.textContent).toBe('Open a diagram to search');
		viewer.destroy();
	});
	it('cleans retained controls after destroy and keeps two instances independent', () => {
		const first = mounted(),
			second = mounted();
		first.query('framework');
		second.query('idea');
		first.viewer.destroy();
		const errors: ErrorEvent[] = [],
			onError = (event: ErrorEvent) => errors.push(event);
		window.addEventListener('error', onError);
		first.next.click();
		first.query('release');
		first.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		window.removeEventListener('error', onError);
		expect(errors).toEqual([]);
		expect(second.input.value).toBe('idea');
		second.next.click();
		expect(second.viewer.controller.state.selectedShape?.id).toBe('s1');
		second.viewer.destroy();
	});
	it('keeps HTML-like text and hostile IDs inert in search results and selection', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.id = '\"]<img src=x onerror=bad>';
		model.pages[0]!.shapes[0]!.text.plainText = '<script>alert(1)</script>';
		const { viewer, root, query, next, status } = mounted();
		viewer.update({ document: model });
		query('<script>');
		next.click();
		expect(status.title).toContain('<script>alert(1)</script>');
		expect(root.querySelector('script,img')).toBeNull();
		expect(root.querySelector('[data-selected="true"]')?.getAttribute('data-shape-id')).toBe(
			model.pages[0]!.shapes[0]!.id,
		);
		viewer.destroy();
	});
	it('allows page-change callbacks to replace the document before stale search selection', () => {
		const { viewer, root, input, query, next } = mounted();
		const selected = vi.fn();
		viewer.update({
			events: {
				'page-change': () => viewer.update({ document: { ...demoDocument } }),
				'shape-select': selected,
			},
		});
		query('framework');
		next.click();
		expect(input.value).toBe('');
		expect(viewer.element.pageIndex).toBe(0);
		expect(root.querySelector('[data-selected="true"]')).toBeNull();
		expect(selected).not.toHaveBeenCalled();
		viewer.destroy();
	});
});
