// @vitest-environment jsdom
/**
 * Unit tests for html2canvas-clone.ts: the shared `onclone` pass swaps each
 * `<pptx-three-view>` in html2canvas's clone for its live pixels plus its DOM
 * overlay, and drops editor-only nodes, before the colour/CSS passes.
 *
 * html2canvas re-creates a custom element as a `<div>` that keeps the
 * element's attributes and gets its own (cloned) shadow root, and it clones
 * into a separate iframe document. Both are simulated here, and the live view
 * is a `<div>` carrying the marker plus the `state`/`canvas`/`flush` surface
 * of the real element (which is not registered in jsdom).
 */
import { describe, expect, it } from 'vitest';

import { THREE_VIEW_MARKER_ATTR } from '../three-view/element';
import { THREE_VIEW_SNAPSHOT_ATTR } from '../three-view/export-snapshot';
import { html2canvasCaptureScope, prepareHtml2CanvasClone } from './html2canvas-clone';

const PIXELS = 'data:image/png;base64,chart';

function liveSlide(): HTMLElement {
	const slide = document.createElement('div');
	const view = document.createElement('div');
	view.setAttribute(THREE_VIEW_MARKER_ATTR, '');
	view.setAttribute('data-state', 'ready');
	const shadow = view.attachShadow({ mode: 'open' });
	shadow.innerHTML =
		'<div class="stage"><canvas></canvas><div class="overlay"><svg class="chrome"><text>Chart Title</text></svg></div></div>';
	const canvas = shadow.querySelector('canvas') as HTMLCanvasElement;
	canvas.width = 40;
	canvas.height = 30;
	(canvas as unknown as { toDataURL: () => string }).toDataURL = () => PIXELS;
	Object.assign(view, { state: 'ready', spec: {}, canvas, flush: () => {} });
	view.append(Object.assign(document.createElement('span'), { className: 'fallback-2d' }));
	const handle = document.createElement('div');
	handle.setAttribute('data-export-ignore', 'true');
	slide.append(view, handle);
	document.body.append(slide);
	return slide;
}

/** What html2canvas hands `onclone`: a copy in its own document, views as shadow-hosting divs. */
function html2canvasClone(live: HTMLElement): { doc: Document; clonedEl: HTMLElement } {
	const doc = document.implementation.createHTMLDocument('capture');
	const clonedEl = doc.importNode(live, true) as HTMLElement;
	doc.body.append(clonedEl);
	const view = clonedEl.querySelector(`[${THREE_VIEW_MARKER_ATTR}]`) as HTMLElement;
	view.attachShadow({ mode: 'open' }).innerHTML = '<canvas></canvas><span>fallback copy</span>';
	return { doc, clonedEl };
}

describe('prepareHtml2CanvasClone', () => {
	it('replaces a ready 3D view with its pixels and overlay, in light and shadow DOM', async () => {
		const live = liveSlide();
		const { doc, clonedEl } = html2canvasClone(live);

		await prepareHtml2CanvasClone(live, doc, clonedEl);

		const view = clonedEl.querySelector(`[${THREE_VIEW_MARKER_ATTR}]`) as HTMLElement;
		for (const root of [view, view.shadowRoot as ShadowRoot]) {
			const img = root.querySelector<HTMLImageElement>(`img[${THREE_VIEW_SNAPSHOT_ATTR}="true"]`);
			expect(img?.getAttribute('src')).toBe(PIXELS);
			const overlay = root.querySelector(`[${THREE_VIEW_SNAPSHOT_ATTR}="overlay"]`);
			expect(overlay?.querySelector('svg.chrome')?.textContent).toBe('Chart Title');
			expect(root.querySelector('canvas')).toBeNull();
			expect(root.querySelector('.fallback-2d')).toBeNull();
		}
		// The copies belong to html2canvas's document, not the live one.
		expect(view.firstElementChild?.ownerDocument).toBe(doc);
		live.remove();
	});

	it('drops editor-only nodes from the clone and leaves the live tree untouched', async () => {
		const live = liveSlide();
		const { doc, clonedEl } = html2canvasClone(live);

		await prepareHtml2CanvasClone(live, doc, clonedEl);

		expect(clonedEl.querySelector('[data-export-ignore="true"]')).toBeNull();
		expect(live.querySelector('[data-export-ignore="true"]')).not.toBeNull();
		expect(live.querySelector('.fallback-2d')).not.toBeNull();
		live.remove();
	});

	it('keeps the 2D fallback of a view that never became ready', async () => {
		const live = liveSlide();
		const liveView = live.querySelector(`[${THREE_VIEW_MARKER_ATTR}]`) as HTMLElement;
		Object.assign(liveView, { state: 'unavailable' });
		const { doc, clonedEl } = html2canvasClone(live);

		await prepareHtml2CanvasClone(live, doc, clonedEl);

		expect(clonedEl.querySelector(`img[${THREE_VIEW_SNAPSHOT_ATTR}]`)).toBeNull();
		expect(clonedEl.querySelector('.fallback-2d')).not.toBeNull();
		live.remove();
	});
});

describe('html2canvasCaptureScope', () => {
	function page() {
		document.head.innerHTML =
			'<style>@font-face { font-family: Deck; src: url(deck.woff2); }</style>';
		document.body.innerHTML = `
			<div class="ribbon"><button>Home</button></div>
			<main class="viewer">
				<aside class="rail"><img alt="slide 1"></aside>
				<div class="viewport"><div class="stage"><div class="shape"><span>Title</span></div><span class="host"></span></div></div>
				<aside class="inspector"><style>.x {}</style><input></aside>
			</main>
			<link rel="stylesheet" href="viewer.css">`;
		const at = <T extends Element>(selector: string) => document.querySelector<T>(selector)!;
		const host = at('.host');
		host.attachShadow({ mode: 'open' }).innerHTML = '<svg><text>3D</text></svg>';
		return { at, host };
	}

	it('keeps the captured branch and the style sources, and skips every other branch', () => {
		const { at, host } = page();
		const skip = html2canvasCaptureScope(at('.stage'));
		for (const kept of [
			'html',
			'head',
			'head style',
			'body',
			'.viewer',
			'.viewport',
			'.stage',
			'.shape',
			'.shape span',
			'link',
			'.inspector style',
		]) {
			expect([kept, skip(at(kept))]).toStrictEqual([kept, false]);
		}
		expect(skip(host.shadowRoot!.querySelector('text')!)).toBeFalsy();
		for (const skipped of [
			'.ribbon',
			'.ribbon button',
			'.rail',
			'.inspector',
			'.inspector input',
		]) {
			expect([skipped, skip(at(skipped))]).toStrictEqual([skipped, true]);
		}
	});

	it('follows a captured element inside a shadow root up through its host', () => {
		const { at, host } = page();
		const inner = host.shadowRoot!.querySelector('svg')!;
		const skip = html2canvasCaptureScope(inner);
		expect(skip(host)).toBeFalsy();
		expect(skip(at('.stage'))).toBeFalsy();
		expect(skip(at('.shape'))).toBeTruthy();
		expect(skip(inner.querySelector('text')!)).toBeFalsy();
	});

	it("still applies the caller's own predicate", () => {
		const { at } = page();
		const skip = html2canvasCaptureScope(at('.stage'), (node) => node.classList.contains('shape'));
		expect(skip(at('.shape'))).toBeTruthy();
		expect(skip(at('.stage'))).toBeFalsy();
	});
});
