import { expect, it } from 'vitest';

it('can be imported outside a worker without installing a message handler', async () => {
	const handler = Reflect.get(globalThis, 'onmessage');
	await import('./parse-worker');
	expect(Reflect.get(globalThis, 'onmessage')).toBe(handler);
});
