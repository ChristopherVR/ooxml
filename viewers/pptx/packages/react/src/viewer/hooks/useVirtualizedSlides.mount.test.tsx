// @vitest-environment happy-dom
/**
 * The rail renders its scroll container only while open, so the container can
 * mount after the hook's first render. The hook must still pick it up: it used
 * a mount-once effect, missed a late container, kept a 0 px viewport and never
 * rendered past the first dozen thumbnails of a long deck.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { useVirtualizedSlides } from './useVirtualizedSlides';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let host: HTMLDivElement;
let root: Root;
let range = { startIndex: 0 };

function Rail({ open }: { open: boolean }): React.ReactElement | null {
	const v = useVirtualizedSlides({ totalItems: 112, itemHeight: 100, overscan: 2 });
	range = { startIndex: v.startIndex };
	return open ? <div data-testid='rail' ref={v.scrollContainerRef} /> : null;
}

function rail(): HTMLDivElement {
	const el = host.querySelector<HTMLDivElement>('[data-testid="rail"]');
	if (!el) {
		throw new Error('rail not mounted');
	}
	Object.defineProperty(el, 'clientHeight', { configurable: true, value: 1000 });
	return el;
}

/** Scroll, then let the frame the window reads scroll in run. */
async function scrollTo(el: HTMLDivElement, top: number): Promise<void> {
	await act(async () => {
		el.scrollTop = top;
		el.dispatchEvent(new Event('scroll'));
		await new Promise<void>((resolve) => {
			requestAnimationFrame(() => resolve());
		});
	});
}

beforeEach(() => {
	host = document.createElement('div');
	document.body.appendChild(host);
	root = createRoot(host);
});

afterEach(() => {
	act(() => root.unmount());
	host.remove();
});

describe('useVirtualizedSlides with a late container', () => {
	it('re-renders once per frame of scrolling, and only when rows change', async () => {
		let renders = 0;
		function Counted(): React.ReactElement {
			const v = useVirtualizedSlides({ totalItems: 112, itemHeight: 100, overscan: 2 });
			renders += 1;
			range = { startIndex: v.startIndex };
			return <div data-testid='rail' ref={v.scrollContainerRef} />;
		}
		act(() => root.render(<Counted />));
		const el = rail();
		const settled = renders;
		await act(async () => {
			for (const top of [3000, 3010, 3020, 3030]) {
				el.scrollTop = top;
				el.dispatchEvent(new Event('scroll'));
			}
			await new Promise<void>((resolve) => {
				requestAnimationFrame(() => resolve());
			});
		});
		expect(renders - settled).toBe(1);
		expect(range.startIndex).toBe(28);
		await scrollTo(el, 3040);
		expect(renders - settled).toBe(1);
	});

	it('tracks a container that mounts after the first render', async () => {
		act(() => root.render(<Rail open={false} />));
		act(() => root.render(<Rail open />));
		await scrollTo(rail(), 5000);
		expect(range.startIndex).toBeGreaterThan(40);
	});

	it('follows a container that is closed and reopened', async () => {
		act(() => root.render(<Rail open />));
		act(() => root.render(<Rail open={false} />));
		act(() => root.render(<Rail open />));
		await scrollTo(rail(), 8000);
		expect(range.startIndex).toBeGreaterThan(70);
	});
});
