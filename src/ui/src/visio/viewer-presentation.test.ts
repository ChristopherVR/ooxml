import { afterEach, expect, it, vi } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { ViewerPresentation, presentationPages } from './viewer-presentation';
import { mapPresentationKey } from './presentation-keys';
import { registerViewerControls } from './office-ui';
import { fixture, shape } from '../../../core/visio/test-fixtures';

const cleanups: (() => void)[] = [];
afterEach(() => {
	for (const cleanup of cleanups.splice(0)) cleanup();
	delete (HTMLElement.prototype as { requestFullscreen?: unknown }).requestFullscreen;
});

async function setup() {
	registerViewerControls();
	const controller = new ViewerController(parseVsdx);
	await controller.load(
		await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${shape('1')}</Shapes>`, attributes: 'BackPage="9"' },
				{ id: '9', contents: `<Shapes>${shape('7')}</Shapes>`, attributes: 'Background="1"' },
				{ id: '1', contents: `<Shapes>${shape('2')}</Shapes>` },
				{ id: '2', contents: `<Shapes>${shape('3')}</Shapes>` },
			],
		}),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const viewport = document.createElement('div');
	viewport.className = 'viewport';
	viewport.tabIndex = 0;
	const status = document.createElement('office-ui-button');
	status.className = 'presentation-mode';
	root.append(viewport, status);
	const presentation = new ViewerPresentation(root, controller);
	const dispose = presentation.wire();
	const unsubscribe = controller.subscribe((state) => presentation.render(state));
	presentation.render(controller.state);
	const key = (init: KeyboardEventInit) =>
		presentation.overlay.dispatchEvent(
			new KeyboardEvent('keydown', { bubbles: true, composed: true, cancelable: true, ...init }),
		);
	const shown = () =>
		presentation.overlay.querySelector('svg')?.getAttribute('aria-label') ?? undefined;
	const indicator = () => presentation.overlay.querySelector('.presentation-page')!.textContent;
	cleanups.push(() => {
		dispose();
		unsubscribe();
		controller.destroy();
		host.remove();
	});
	return { controller, root, host, viewport, status, presentation, key, shown, indicator };
}

it('presents foreground pages only and composes the background page under them', async () => {
	const { controller, presentation, shown, indicator } = await setup();
	expect(presentationPages(controller.state.document).map((page) => page.id)).toEqual([
		'0',
		'1',
		'2',
	]);
	presentation.start();
	expect(presentation.active).toBe(true);
	expect(presentation.overlay.hidden).toBe(false);
	expect(shown()).toBe('Page 1');
	expect(indicator()).toBe('Page 1 (1 of 3)');
	// Background shape 7 is drawn under the foreground page; nothing is selectable.
	const svg = presentation.overlay.querySelector('svg')!;
	expect(svg.querySelectorAll('g').length).toBeGreaterThan(1);
	expect(svg.querySelector('[data-shape-id]')).toBeNull();
});

it('navigates with Visio presentation keys, clicks and typed page numbers', async () => {
	const { presentation, key, shown, indicator } = await setup();
	presentation.start();
	key({ key: 'ArrowRight' });
	expect(shown()).toBe('Page 3');
	key({ key: ' ' });
	expect(shown()).toBe('Page 4');
	key({ key: 'PageDown' });
	expect(shown()).toBe('Page 4');
	key({ key: 'Backspace' });
	expect(shown()).toBe('Page 3');
	key({ key: 'Home' });
	expect(shown()).toBe('Page 1');
	key({ key: 'End' });
	expect(shown()).toBe('Page 4');
	key({ key: 'ArrowUp' });
	key({ key: 'PageUp' });
	expect(shown()).toBe('Page 1');
	key({ key: '3' });
	expect(indicator()).toBe('Go to page 3');
	key({ key: 'Enter' });
	expect(indicator()).toBe('Page 4 (3 of 3)');
	key({ key: '9' });
	key({ key: 'Enter' });
	expect(indicator()).toBe('No page 9 (3 pages)');
	expect(shown()).toBe('Page 4');
	key({ key: 'Home' });
	presentation.overlay.querySelector<HTMLElement>('.presentation-stage')!.click();
	expect(shown()).toBe('Page 3');
	presentation.overlay.querySelector<HTMLButtonElement>('[data-present="previous"]')!.click();
	expect(shown()).toBe('Page 1');
	expect(
		presentation.overlay.querySelector<HTMLButtonElement>('[data-present="previous"]')!.disabled,
	).toBe(true);
});

it('keeps editing shortcuts away from the drawing and exits on Escape restoring the view', async () => {
	const { controller, root, viewport, presentation, key } = await setup();
	controller.setPage(2);
	controller.setZoom(1.5);
	controller.selectShape({ id: '2', name: 'Sheet.2', pageId: '1' });
	const before = controller.state;
	viewport.focus();
	const seen = vi.fn();
	root.addEventListener('keydown', seen);
	presentation.start();
	expect(presentation.overlay.textContent).toContain('(2 of 3)');
	expect(viewport.closest('[inert]') ?? viewport.hasAttribute('inert')).toBeTruthy();
	key({ key: 'ArrowRight' });
	key({ key: 'Delete' });
	key({ key: 'z', ctrlKey: true });
	expect(seen).not.toHaveBeenCalled();
	key({ key: 'Escape' });
	expect(presentation.active).toBe(false);
	expect(presentation.overlay.hidden).toBe(true);
	expect(presentation.overlay.querySelector('svg')).toBeNull();
	expect(viewport.hasAttribute('inert')).toBe(false);
	expect(root.activeElement).toBe(viewport);
	expect(controller.state).toBe(before);
	expect(controller.state.pageIndex).toBe(2);
	expect(controller.state.zoom).toBe(1.5);
	expect(controller.state.selectedShapes.map((item) => item.id)).toEqual(['2']);
});

it('falls back to the in-element overlay without the Fullscreen API', async () => {
	const { presentation } = await setup();
	expect('requestFullscreen' in presentation.overlay).toBe(false);
	presentation.start();
	expect(presentation.overlay.dataset.fallback).toBe('true');
	presentation.overlay.querySelector<HTMLButtonElement>('[data-present="exit"]')!.click();
	expect(presentation.active).toBe(false);
	expect(presentation.overlay.dataset.fallback).toBeUndefined();
});

it('uses the Fullscreen API when present and keeps the overlay when it is refused', async () => {
	const request = vi.fn(() => Promise.reject(new Error('denied')));
	(HTMLElement.prototype as { requestFullscreen?: unknown }).requestFullscreen = request;
	const { presentation } = await setup();
	presentation.start();
	expect(request).toHaveBeenCalledOnce();
	await Promise.resolve();
	expect(presentation.active).toBe(true);
	expect(presentation.overlay.dataset.fallback).toBe('true');
});

it('starts from the status bar button, follows edits and disables without foreground pages', async () => {
	const { controller, status, presentation, shown } = await setup();
	status.click();
	expect(presentation.active).toBe(true);
	presentation.go(2);
	const document = structuredClone(controller.state.document!);
	document.pages.splice(2, 1);
	controller.setDocument(document);
	expect(shown()).toBe('Page 4');
	controller.setDocument(null);
	expect(presentation.active).toBe(false);
	expect((status as HTMLElement & { disabled: boolean }).disabled).toBe(true);
});

it('maps keys like Visio and ignores chords', () => {
	const buffer = { digits: '' };
	expect(mapPresentationKey({ key: 'Enter' }, buffer)).toEqual({ type: 'next' });
	expect(mapPresentationKey({ key: 'ArrowRight', ctrlKey: true }, buffer)).toEqual({
		type: 'none',
	});
	mapPresentationKey({ key: '1' }, buffer);
	mapPresentationKey({ key: 'Shift' }, buffer);
	mapPresentationKey({ key: '2' }, buffer);
	expect(mapPresentationKey({ key: 'Enter' }, buffer)).toEqual({ type: 'goto', page: 12 });
	mapPresentationKey({ key: '0' }, buffer);
	expect(mapPresentationKey({ key: 'Enter' }, buffer)).toEqual({ type: 'none' });
	expect(mapPresentationKey({ key: 'Escape' }, buffer)).toEqual({ type: 'exit' });
});
