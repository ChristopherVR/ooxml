import { expect, it } from 'vitest';
import { crc32 } from './crc32';
it('retains standard CRC32 and incremental raw-state semantics', () => {
	const bytes = new TextEncoder().encode('123456789');
	expect((crc32(bytes) ^ 0xffffffff) >>> 0).toBe(0xcbf43926);
	expect(crc32(bytes.subarray(4), crc32(bytes.subarray(0, 4)))).toBe(crc32(bytes));
});
