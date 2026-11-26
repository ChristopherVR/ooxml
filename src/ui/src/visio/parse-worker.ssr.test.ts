import { expect, it } from 'vitest';

it('can import all Visio worker entries on the server without installing a handler', async () => {
	const handler = Reflect.get(globalThis, 'onmessage');
	await import('./parse-worker');
	await import('./edit-worker');
	await import('./clipboard-worker');
	expect(Reflect.get(globalThis, 'onmessage')).toBe(handler);
});
