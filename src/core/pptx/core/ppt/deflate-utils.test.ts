import { deflateSync } from 'node:zlib';

import { describe, expect, it } from 'vitest';

import { inflateZlib } from './deflate-utils';

describe('inflateZlib', () => {
	it('inflates exactly the bytes of a Node Buffer view, not its backing store', async () => {
		const compressed = deflateSync(Buffer.from('hello, hello, hello'));
		// A pooled or sliced read hands back a view like this, at a non-zero offset.
		const store = new Uint8Array(compressed.length + 64);
		store.set(compressed, 32);
		const inflated = await inflateZlib(Buffer.from(store.buffer, 32, compressed.length));
		expect(inflated && new TextDecoder().decode(inflated)).toBe('hello, hello, hello');
	});
});
