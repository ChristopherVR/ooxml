// @vitest-environment happy-dom
/**
 * `useViewportFit` measures the viewport before paint.
 *
 * React flushes a state update made in a layout effect before the browser paints, and one made
 * in a passive effect (`useEffect`) only after: the browser can paint the commit in between,
 * which showed the slide at its authored size for a frame before it jumped to the fitted size.
 * A parent's layout effect runs after its children's in the same commit, so the probe below
 * records whether the viewport had been measured by the end of that commit's layout phase.
 */
import { createElement, useCallback, useLayoutEffect, useRef } from 'react';
import { flushSync } from 'react-dom';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { useViewportFit } from './useViewportFit';
import { useViewportNode } from './useViewportNode';

const canvasSize = { width: 1280, height: 720 };

/** Layout reads of the viewport so far. */
let layoutReads = 0;

/** A viewport laid out at `width` x `height` whole pixels. */
function layOut(node: HTMLDivElement, width: number, height: number): void {
	const read = (value: number) => () => {
		layoutReads += 1;
		return value;
	};
	Object.defineProperties(node, {
		clientWidth: { configurable: true, get: read(width) },
		clientHeight: { configurable: true, get: read(height) },
	});
}

/** Records how many viewport reads had happened when the mounting commit's layout phase ended. */
function LayoutPhaseProbe(props: {
	width: number;
	height: number;
	onLayoutEnd: (reads: number) => void;
}) {
	const { onLayoutEnd } = props;
	useLayoutEffect(() => onLayoutEnd(layoutReads), [onLayoutEnd]);
	return createElement(Viewer, { width: props.width, height: props.height });
}

function Slide(props: { fitScale: number }) {
	return createElement('div', {
		'data-testid': 'slide',
		style: {
			width: canvasSize.width * props.fitScale,
			height: canvasSize.height * props.fitScale,
		},
	});
}

function Viewer(props: { width: number; height: number }) {
	const { viewportRef, viewportNode, setCanvasViewportNode } = useViewportNode();
	const { fitScale } = useViewportFit(
		canvasSize,
		viewportRef,
		{ fitPadding: 0, maxFitScale: null },
		viewportNode,
	);
	const { width, height } = props;
	const attach = useCallback(
		(node: HTMLDivElement | null) => {
			if (node) {
				layOut(node, width, height);
			}
			setCanvasViewportNode(node);
		},
		[setCanvasViewportNode, width, height],
	);
	return createElement(
		'div',
		{ ref: attach, 'data-testid': 'viewport' },
		createElement(Slide, { fitScale }),
	);
}

/** An object-ref caller whose ref is assigned only after the mounting commit. */
function LateObjectRefViewer(props: { onRef: (assign: () => void) => void }) {
	const viewportRef = useRef<HTMLDivElement | null>(null);
	const { fitScale } = useViewportFit(canvasSize, viewportRef, {
		fitPadding: 0,
		maxFitScale: null,
	});
	const { onRef } = props;
	useLayoutEffect(() => {
		onRef(() => {
			const node = document.createElement('div');
			layOut(node, 640, 360);
			viewportRef.current = node;
		});
	}, [onRef]);
	return createElement(Slide, { fitScale });
}

describe('useViewportFit', () => {
	let host: HTMLDivElement;
	let root: Root;

	beforeEach(() => {
		// `flushSync` outside `act` on purpose: `act` would also flush the passive effects.
		vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', false);
		vi.stubGlobal(
			'ResizeObserver',
			class {
				observe() {}
				disconnect() {}
			},
		);
		layoutReads = 0;
		host = document.createElement('div');
		document.body.append(host);
		root = createRoot(host);
	});

	afterEach(() => {
		flushSync(() => root.unmount());
		host.remove();
		vi.unstubAllGlobals();
	});

	const slideSize = () => {
		const slide = host.querySelector<HTMLElement>('[data-testid="slide"]')!;
		return {
			width: Number.parseFloat(slide.style.width),
			height: Number.parseFloat(slide.style.height),
		};
	};

	it('measures the viewport in the layout phase of the commit that mounts it', () => {
		let readsAtLayoutEnd = -1;
		flushSync(() =>
			root.render(
				createElement(LayoutPhaseProbe, {
					width: 640,
					height: 360,
					onLayoutEnd: (reads: number) => {
						readsAtLayoutEnd = reads;
					},
				}),
			),
		);
		// Measured before paint, so the fitted size is what the first painted frame shows.
		expect(readsAtLayoutEnd).toBeGreaterThan(0);
		expect(slideSize()).toStrictEqual({ width: 640, height: 360 });
	});

	it('still retries on the next frame when an object ref is assigned after the commit', () => {
		const frames: FrameRequestCallback[] = [];
		vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
			frames.push(callback);
			return frames.length;
		});
		vi.stubGlobal('cancelAnimationFrame', () => undefined);
		let assign: () => void = () => undefined;
		flushSync(() =>
			root.render(
				createElement(LateObjectRefViewer, {
					onRef: (next: () => void) => {
						assign = next;
					},
				}),
			),
		);
		// No node yet: the authored size, and a retry queued.
		expect(slideSize()).toStrictEqual({ width: 1280, height: 720 });
		expect(frames).toHaveLength(1);
		assign();
		flushSync(() => frames.shift()!(0));
		expect(slideSize()).toStrictEqual({ width: 640, height: 360 });
	});
});
