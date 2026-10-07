import { describe, expect, it } from 'vitest';

import { resolveSourceToBuffer } from './source';

describe('resolveSourceToBuffer', () => {
	it('normalises Uint8Array views to their exact byte range', async () => {
		const backing = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7]);
		const view = new Uint8Array(backing.buffer, 2, 3);
		const buffer = await resolveSourceToBuffer(view);
		expect(Array.from(new Uint8Array(buffer))).toStrictEqual([2, 3, 4]);
	});

	it('passes ArrayBuffers through and reads Blobs', async () => {
		const raw = new Uint8Array([9, 8, 7]).buffer;
		await expect(resolveSourceToBuffer(raw)).resolves.toBe(raw);

		const blob = new Blob([new Uint8Array([1, 2])]);
		const fromBlob = await resolveSourceToBuffer(blob);
		expect(new Uint8Array(fromBlob)).toHaveLength(2);
	});
});
