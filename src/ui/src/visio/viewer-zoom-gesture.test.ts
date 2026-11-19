// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { mountViewer } from './binding';
import { scrollRange, settledOrigin, type ZoomLayout } from './viewer-zoom-gesture';

afterEach(() => {
	vi.useRealTimers();
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

/** A mounted viewer whose drawing window has a layout: jsdom lays nothing out. */
function laidOut() {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	const root = viewer.element.shadowRoot!;
	const viewport = root.querySelector<HTMLElement>('.viewport')!;
	const paper = () => root.querySelector<SVGSVGElement>('svg.paper')!;
	for (const [name, value] of [
		['clientWidth', 800],
		['clientHeight', 600],
	] as const)
		Object.defineProperty(viewport, name, { configurable: true, value });
	// The paper sits 20 px into the scrolled content; only its laid-out size counts, as in a
	// browser, where the preview transform is not layout.
	vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
		if (this === viewport) return new DOMRect(0, 0, 800, 600);
		if (this instanceof SVGSVGElement && this.classList.contains('paper')) {
			const style = (this as SVGSVGElement).style;
			return new DOMRect(
				20 - viewport.scrollLeft,
				20 - viewport.scrollTop,
				Number.parseFloat(style.width),
				Number.parseFloat(style.height),
			);
		}
		return new DOMRect();
	});
	vi.useFakeTimers({
		toFake: ['setTimeout', 'clearTimeout', 'requestAnimationFrame', 'cancelAnimationFrame'],
	});
	const page = demoDocument.pages[0]!;
	const width = (zoom: number) => `${page.width * 96 * zoom}px`;
	const wheel = (deltaY: number) =>
		viewport.dispatchEvent(
			new WheelEvent('wheel', { deltaY, ctrlKey: true, bubbles: true, cancelable: true }),
		);
	const touch = (type: string, id: number, x: number, y: number) =>
		viewport.dispatchEvent(
			Object.assign(
				new MouseEvent(type, { clientX: x, clientY: y, bubbles: true, cancelable: true }),
				{
					pointerType: 'touch',
					pointerId: id,
				},
			),
		);
	return { viewer, root, viewport, paper, width, wheel, touch };
}

describe('interactive zoom', () => {
	it('scales the drawing while Ctrl+wheel turns and renders it once when the wheel stops', () => {
		const { viewer, root, viewport, paper, width, wheel } = laidOut();
		viewport.scrollLeft = 30;
		viewport.scrollTop = 12;
		wheel(-120);
		wheel(-120);
		wheel(-120);
		vi.advanceTimersByTime(16);
		// The page keeps its rendered size; the user sees it scaled to the new zoom.
		expect(viewer.controller.state.zoom).toBe(1);
		expect(paper().style.width).toBe(width(1));
		expect(paper().style.transform).toMatch(/ scale\(1\.331\)$/);
		expect(
			root.querySelector<HTMLElement & { value: number }>('office-ui-zoom-slider')!.value,
		).toBe(133);
		vi.advanceTimersByTime(400);
		// The same end zoom and scroll position as zooming step by step.
		const zoom = 1 * 1.1 * 1.1 * 1.1;
		expect(viewer.controller.state.zoom).toBe(zoom);
		expect(Number.parseFloat(paper().style.width)).toBeCloseTo(Number.parseFloat(width(zoom)), 1);
		expect(paper().style.transform).toBe('');
		expect([viewport.scrollLeft, viewport.scrollTop]).toEqual([30, 12]);
		viewer.destroy();
	});

	it('keeps the pinch point under the fingers once the pinch ends', () => {
		const { viewer, paper, width, touch } = laidOut();
		const page = demoDocument.pages[0]!;
		touch('pointerdown', 1, 100, 100);
		touch('pointerdown', 2, 200, 100);
		touch('pointermove', 2, 300, 100);
		vi.advanceTimersByTime(16);
		expect(viewer.controller.state.zoom).toBe(1);
		expect(paper().style.transform).toContain('scale(2)');
		touch('pointerup', 2, 300, 100);
		touch('pointerup', 1, 100, 100);
		expect(viewer.controller.state.zoom).toBe(2);
		expect(paper().style.width).toBe(width(2));
		// The page point under the pinch start (150, 100) is now under the fingers' midpoint.
		const box = paper().getBoundingClientRect();
		const fraction = (150 - 20) / (page.width * 96);
		expect(box.left + fraction * box.width).toBeCloseTo(200);
		viewer.destroy();
	});

	it('gives way to a zoom set by a command or the host, and settles before pointer input', () => {
		const { viewer, viewport, paper, width, wheel } = laidOut();
		wheel(-120);
		vi.advanceTimersByTime(16);
		viewer.element.zoom = 2;
		expect(paper().style.transform).toBe('');
		vi.advanceTimersByTime(400);
		expect(viewer.controller.state.zoom).toBe(2);
		wheel(120);
		viewport.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
		expect(viewer.controller.state.zoom).toBe(2 / 1.1);
		expect(Number.parseFloat(paper().style.width)).toBeCloseTo(
			Number.parseFloat(width(2 / 1.1)),
			1,
		);
		viewer.destroy();
	});

	it('gives a press on a handle the settled render replaced to the handle now under the pointer', () => {
		const { viewer, root, viewport, paper, wheel } = laidOut();
		// A selection handle: drawn again, as a new element, whenever the zoom changes.
		const draw = () =>
			paper().appendChild(document.createElementNS('http://www.w3.org/2000/svg', 'circle'));
		let handle = draw();
		viewer.controller.subscribe(() => {
			handle.remove();
			handle = draw();
		});
		Object.defineProperty(root, 'elementFromPoint', { configurable: true, value: () => handle });
		const presses: [EventTarget | null, number][] = [];
		viewport.addEventListener(
			'pointerdown',
			(event) => presses.push([event.target, event.pointerId]),
			{ capture: true },
		);
		wheel(-120);
		vi.advanceTimersByTime(16);
		const pressed = handle;
		pressed.dispatchEvent(
			new PointerEvent('pointerdown', { bubbles: true, composed: true, pointerId: 3 }),
		);
		expect(viewer.controller.state.zoom).toBeCloseTo(1.1);
		expect(handle).not.toBe(pressed);
		expect(presses).toEqual([[handle, 3]]);
		viewer.destroy();
	});
});

describe('zoom preview geometry', () => {
	const layout: ZoomLayout = {
		origin: { x: 0, y: 0 },
		client: { x: 800, y: 600 },
		padding: { left: 4, right: 4, top: 16, bottom: 16 },
		page: { x: 1000, y: 500 },
	};
	it('measures the scroll range from the paper and the padding', () => {
		expect(scrollRange(layout, 1)).toEqual({ x: 208, y: 0 });
		expect(scrollRange(layout, 2)).toEqual({ x: 1208, y: 432 });
	});
	it('centres a small page and keeps a large one at the clamped scroll offsets', () => {
		// Half size: centred in the padding box.
		expect(settledOrigin(layout, { x: 0, y: 0 }, 0.5, 0.5, { kind: 'scroll' })).toEqual({
			x: 4 + (792 - 500) / 2,
			y: 16 + (568 - 250) / 2,
		});
		// Zooming out to 0.9 and back to 2 clamps the scroll offsets as each step did.
		expect(settledOrigin(layout, { x: 500, y: 300 }, 2, 0.9, { kind: 'scroll' })).toEqual({
			x: 4 - 108,
			y: 16 - 0,
		});
	});
});
