import { afterEach, describe, expect, it, vi } from 'vitest';

import { fitScale } from './state/navigation';
import { viewportClientSize } from './viewport-client-size';

describe('viewportClientSize action', () => {
	afterEach(() => {
		vi.unstubAllGlobals();
	});

	it('reports the fractional size, so the fitted slide never exceeds the viewport', () => {
		let notify: () => void = () => undefined;
		let disconnected = false;
		vi.stubGlobal(
			'ResizeObserver',
			class {
				constructor(callback: () => void) {
					notify = callback;
				}
				observe() {}
				disconnect() {
					disconnected = true;
				}
			},
		);
		let size = { width: 801.5, height: 700 };
		const node = document.createElement('div');
		Object.defineProperties(node, {
			clientWidth: { get: () => Math.round(size.width) },
			clientHeight: { get: () => Math.round(size.height) },
			offsetWidth: { get: () => Math.round(size.width) },
			offsetHeight: { get: () => Math.round(size.height) },
		});
		node.getBoundingClientRect = () => new DOMRect(0, 0, size.width, size.height);
		const reports: { width: number; height: number }[] = [];
		const action = viewportClientSize(node, (next) => reports.push(next));
		// Synchronous on mount, as `bind:clientWidth` was, so the first frame is already fitted.
		expect(reports).toStrictEqual([{ width: 801.5, height: 700 }]);
		const [first] = reports;
		const scale = fitScale(first!.width, first!.height, 960, 540, 0, null);
		expect(960 * scale).toBeLessThanOrEqual(801.5 + 1e-9);
		size = { width: 1200, height: 451.5 };
		const later: { width: number; height: number }[] = [];
		action.update((next) => later.push(next));
		notify();
		expect(later).toStrictEqual([{ width: 1200, height: 451.5 }]);
		action.destroy();
		expect(disconnected).toBeTruthy();
	});

	it('also watches the border box, so a padding-only change is reported', () => {
		// A content-box viewport with a fixed width: its padding changes the client and border
		// boxes but not the content box a default ResizeObserver (and `bind:clientWidth`'s
		// replacement) watches. Svelte's own binding observed the border box.
		const observers: { box: string | undefined; callback: () => void }[] = [];
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
		let size = { width: 849.5, height: 748 };
		const node = document.createElement('div');
		Object.defineProperties(node, {
			clientWidth: { get: () => Math.round(size.width) },
			clientHeight: { get: () => Math.round(size.height) },
			offsetWidth: { get: () => Math.round(size.width) },
			offsetHeight: { get: () => Math.round(size.height) },
		});
		node.getBoundingClientRect = () => new DOMRect(0, 0, size.width, size.height);
		const reports: { width: number; height: number }[] = [];
		const action = viewportClientSize(node, (next) => reports.push(next));
		size = { width: 801.5, height: 700 };
		const borderBox = observers.find((observer) => observer.box === 'border-box');
		expect(borderBox).toBeDefined();
		borderBox!.callback();
		expect(reports.at(-1)).toStrictEqual({ width: 801.5, height: 700 });
		action.destroy();
	});
});
