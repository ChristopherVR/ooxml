import { expect, test } from 'vitest';

import { isSelfContainedGlb } from './glb-import';

function model(json: string): Uint8Array {
	const source = new TextEncoder().encode(json.padEnd(Math.ceil(json.length / 4) * 4, ' '));
	const result = new Uint8Array(20 + source.length);
	const header = new DataView(result.buffer);
	header.setUint32(0, 0x46546c67, true);
	header.setUint32(4, 2, true);
	header.setUint32(8, result.length, true);
	header.setUint32(12, source.length, true);
	header.setUint32(16, 0x4e4f534a, true);
	result.set(source, 20);
	return result;
}

test('admits GLB 2.0 payloads without external references', () => {
	const valid = model('{"asset":{"version":"2.0"},"buffers":[{"byteLength":16}]}');
	expect(isSelfContainedGlb(valid)).toBe(true);
	expect(isSelfContainedGlb(valid.subarray(0, -1))).toBe(false);
	expect(
		isSelfContainedGlb(
			model(
				'{"asset":{"version":"2.0"},"images":[{"\\u0075ri":"https://example.invalid/texture.png"}]}',
			),
		),
	).toBe(false);
});
