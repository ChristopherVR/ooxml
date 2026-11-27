// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import type { FormatDialog } from './font-dialog';
import { lazyDialog } from './lazy-dialog';

function fake(): FormatDialog & { opened: number } {
	let open = false;
	return {
		opened: 0,
		element: document.createElement('div'),
		open() {
			open = true;
			this.opened++;
		},
		close() {
			open = false;
		},
		setLocale() {},
		get isOpen() {
			return open;
		},
	};
}

/** A load the test settles by hand. */
function deferred() {
	let settle!: (value: FormatDialog) => void;
	let fail!: (error: Error) => void;
	const promise = new Promise<FormatDialog>((resolve, reject) => {
		settle = resolve;
		fail = reject;
	});
	return { promise, settle, fail };
}

describe('lazyDialog', () => {
	it('counts as open while it loads and opens once, however often it is asked', async () => {
		const load = deferred();
		const dialog = lazyDialog(() => load.promise);
		dialog.open();
		dialog.open();
		expect(dialog.isOpen).toBe(true);
		const real = fake();
		load.settle(real);
		await load.promise;
		await Promise.resolve();
		expect(real.opened).toBe(1);
		expect(dialog.isOpen).toBe(true);
	});

	it('does not open after it was closed during the load', async () => {
		const load = deferred();
		const dialog = lazyDialog(() => load.promise);
		dialog.open();
		dialog.close();
		expect(dialog.isOpen).toBe(false);
		const real = fake();
		load.settle(real);
		await load.promise;
		await Promise.resolve();
		expect(real.opened).toBe(0);
	});

	it('reports a failed load and tries again on the next open', async () => {
		const failed = vi.fn();
		let attempt = 0;
		const real = fake();
		const dialog = lazyDialog(
			() => (attempt++ ? Promise.resolve(real) : Promise.reject(new Error('offline'))),
			failed,
		);
		dialog.open();
		await new Promise((resolve) => setTimeout(resolve));
		expect(failed).toHaveBeenCalledWith(expect.objectContaining({ message: 'offline' }));
		expect(dialog.isOpen).toBe(false);
		dialog.open();
		await new Promise((resolve) => setTimeout(resolve));
		expect(real.opened).toBe(1);
	});
});
