// @vitest-environment node
import { expect, it, vi } from 'vitest';

it('can import all Visio worker entries on the server without installing a handler', async () => {
	// Bun already exposes both globals on its main thread. Give Node the same shape.
	if (typeof self === 'undefined') vi.stubGlobal('self', globalThis);
	if (typeof self.postMessage !== 'function') vi.stubGlobal('postMessage', () => {});
	const handler = Reflect.get(globalThis, 'onmessage');
	try {
		for (const load of [
			() => import('./parse-worker'),
			() => import('./edit-worker'),
			() => import('./clipboard-worker'),
		]) {
			await load();
			expect(Reflect.get(globalThis, 'onmessage')).toBe(handler);
		}
	} finally {
		vi.unstubAllGlobals();
		if (handler === undefined) Reflect.deleteProperty(globalThis, 'onmessage');
		else Reflect.set(globalThis, 'onmessage', handler);
	}
});
