// @vitest-environment happy-dom
/**
 * `useViewportFit` fits to the viewport's fractional client size, and follows every change of
 * it, padding-only ones included.
 */
import { act, createElement, useCallback } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { useViewportFit } from './useViewportFit';
import { useViewportNode } from './useViewportNode';

const canvasSize = { width: 1280, height: 720 };

/** The size the viewport is laid out at; the stubs below read it live. */
let layout = { width: 0, height: 0 };

/** What a browser reports for a box laid out at a fractional size. */
function layOut(node: HTMLDivElement): void {
	Object.defineProperties(node, {
		clientWidth: { configurable: true, get: () => Math.round(layout.width) },
		clientHeight: { configurable: true, get: () => Math.round(layout.height) },
		offsetWidth: { configurable: true, get: () => Math.round(layout.width) },
		offsetHeight: { configurable: true, get: () => Math.round(layout.height) },
	});
	node.getBoundingClientRect = () => new DOMRect(0, 0, layout.width, layout.height);
}

function Viewer() {
	const { viewportRef, viewportNode, setCanvasViewportNode } = useViewportNode();
	const { fitScale } = useViewportFit(
		canvasSize,
		viewportRef,
		{ fitPadding: 0, maxFitScale: null },
		viewportNode,
	);
	const attach = useCallback(
		(node: HTMLDivElement | null) => {
			if (node) {
				layOut(node);
			}
			setCanvasViewportNode(node);
		},
		[setCanvasViewportNode],
	);
	return createElement(
		'div',
		{ ref: attach },
		createElement('div', {
			'data-testid': 'slide',
			style: { width: canvasSize.width * fitScale, height: canvasSize.height * fitScale },
		}),
	);
}

describe('useViewportFit at fractional sizes', () => {
	let host: HTMLDivElement;
	let root: Root;
	let observers: { box: string | undefined; callback: () => void }[];

	beforeEach(() => {
		vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
		observers = [];
		vi.stubGlobal(
			'ResizeObserver',
			class {
				private readonly callback: () => void;
				constructor(callback: () => void) {
					this.callback = callback;
				}
				observe(_node: Element, options?: ResizeObserverOptions) {
					observers.push({ box: options?.box, callback: this.callback });
				}
				disconnect() {}
			},
		);
		host = document.createElement('div');
		document.body.append(host);
		root = createRoot(host);
	});

	afterEach(() => {
		act(() => root.unmount());
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

	it.each([
		[801.5, 700, 801.5],
		[1000.6, 800, 1000.6],
		[1400, 451.5, 802.6666666666666],
	])('fits a %spx x %spx viewport without exceeding it', (width, height, expectedSlideWidth) => {
		layout = { width, height };
		act(() => root.render(createElement(Viewer)));
		const slide = slideSize();
		expect(slide.width).toBeCloseTo(expectedSlideWidth, 6);
		expect(slide.width).toBeLessThanOrEqual(width + 1e-9);
		expect(slide.height).toBeLessThanOrEqual(height + 1e-9);
	});

	it('refits on a padding-only change, which only the border box shows', () => {
		// A content-box viewport losing its padding: the client size drops, the content box a
		// default ResizeObserver watches stays put.
		layout = { width: 849.5, height: 748 };
		act(() => root.render(createElement(Viewer)));
		expect(slideSize().width).toBeCloseTo(849.5, 6);
		layout = { width: 801.5, height: 700 };
		const borderBox = observers.find((observer) => observer.box === 'border-box');
		expect(borderBox).toBeDefined();
		act(() => borderBox!.callback());
		expect(slideSize().width).toBeCloseTo(801.5, 6);
	});
});
