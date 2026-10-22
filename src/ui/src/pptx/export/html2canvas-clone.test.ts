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
import {
	html2canvasCaptureScope,
	pinCapturedSize,
	prepareHtml2CanvasClone,
} from './html2canvas-clone';

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

/** The elements html2canvas-pro would clone: it visits a child only when its parent is kept. */
function cloned(skip: (node: Element) => boolean): Set<Element> {
	const kept = new Set<Element>();
	const visit = (node: Element) => {
		kept.add(node);
		const children = [...(node.shadowRoot?.children ?? []), ...node.children].flatMap((child) =>
			child.localName === 'slot'
				? [child, ...(child as HTMLSlotElement).assignedElements()]
				: [child],
		);
		for (const child of children) if (!skip(child)) visit(child);
	};
	visit(document.documentElement);
	return kept;
}

describe('html2canvasCaptureScope', () => {
	function page() {
		document.head.innerHTML =
			'<style>@font-face { font-family: Deck; src: url(deck.woff2); }</style>';
		document.body.innerHTML = `
			<div class="ribbon"><button>Home</button></div>
			<main class="viewer">
				<aside class="rail"><img alt="slide 1"></aside>
				<div class="viewport"><div class="stage"><div class="shape"><a href="https://example.com/"><span>Linked</span></a></div><span class="host"></span></div></div>
				<aside class="inspector"><style>.x {}</style><input></aside>
			</main>
			<link rel="stylesheet" href="viewer.css">`;
		const at = <T extends Element>(selector: string) => document.querySelector<T>(selector)!;
		const host = at('.host');
		host.attachShadow({ mode: 'open' }).innerHTML = '<svg><text>3D</text></svg>';
		return { at, host };
	}

	it("clones the captured branch, shells of its ancestors' other children and the style sheets", () => {
		const { at, host } = page();
		const kept = cloned(html2canvasCaptureScope(at('.stage')));
		// Shells keep sibling selectors (.rail + .viewport) and their own sizes; contents are skipped.
		for (const selector of [
			'head style',
			'.ribbon',
			'.viewer',
			'.rail',
			'.viewport',
			'.stage',
			'.shape a span',
			'body > link',
			'.inspector',
			'.inspector style',
		]) {
			expect([selector, kept.has(at(selector))]).toStrictEqual([selector, true]);
		}
		expect(kept.has(host.shadowRoot!.querySelector('text')!)).toBeTruthy();
		for (const selector of ['.ribbon button', '.rail img', '.inspector input']) {
			expect([selector, kept.has(at(selector))]).toStrictEqual([selector, false]);
		}
	});

	it('keeps a style sheet a slot shows, with the wrapper around that slot', () => {
		const { at } = page();
		const host = document.createElement('div');
		host.attachShadow({ mode: 'open' }).innerHTML =
			'<div class="css"><div class="deep"><slot name="css"></slot></div></div><div class="content"><slot></slot></div>';
		host.innerHTML = '<style slot="css">.x {}</style><div class="target"></div>';
		at('.viewport').append(host);
		const kept = cloned(html2canvasCaptureScope(host.querySelector('.target')!));
		expect(kept.has(host.shadowRoot!.querySelector('.css')!)).toBeTruthy();
		expect(kept.has(host.shadowRoot!.querySelector('.deep')!)).toBeTruthy();
	});

	it('walks up through a link around the captured element', () => {
		const { at } = page();
		const linked = at('.shape a span');
		expect(cloned(html2canvasCaptureScope(linked)).has(linked)).toBeTruthy();
	});

	it('follows a captured element inside a shadow root and through a slot', () => {
		const { at, host } = page();
		const inner = host.shadowRoot!.querySelector('svg')!;
		const kept = cloned(html2canvasCaptureScope(inner));
		expect([kept.has(host), kept.has(inner), kept.has(at('.shape a'))]).toStrictEqual([
			true,
			true,
			false,
		]);

		const framed = document.createElement('div');
		framed.attachShadow({ mode: 'open' }).innerHTML = '<div class="frame"><slot></slot></div>';
		const slotted = document.createElement('div');
		framed.append(slotted);
		at('.viewport').append(framed);
		// The slot's wrapper inside the shadow root is what places it: without it, nothing renders.
		const slottedKept = cloned(html2canvasCaptureScope(slotted));
		expect(slottedKept.has(framed.shadowRoot!.querySelector('.frame')!)).toBeTruthy();
		expect(slottedKept.has(slotted)).toBeTruthy();
	});

	it("still applies the caller's own predicate", () => {
		const { at } = page();
		const skip = html2canvasCaptureScope(at('.stage'), (node) => node.classList.contains('shape'));
		expect(skip(at('.shape'))).toBeTruthy();
		expect(skip(at('.stage'))).toBeFalsy();
	});
});

describe('pinCapturedSize', () => {
	it('keeps the clone at the live layout width only, and leaves an element with no box alone', () => {
		const original = document.createElement('div');
		Object.defineProperties(original, {
			offsetWidth: { value: 960 },
			offsetHeight: { value: 540 },
		});
		document.body.append(original);
		const clone = document.createElement('div');
		pinCapturedSize(original, clone);
		// No height: a fixed one would change how percentage heights inside resolve.
		expect([clone.style.width, clone.style.maxWidth, clone.style.height]).toStrictEqual([
			'960px',
			'none',
			'',
		]);
		const unsized = document.createElement('div');
		pinCapturedSize(document.createElement('div'), unsized);
		expect(unsized.getAttribute('style')).toBeNull();
	});
});
