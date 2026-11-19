import { describe, expect, it } from 'vitest';
import { defaultXmlHeapBytes, runtimeHeapLimit } from './heap-limit';

const MiB = 1024 * 1024;
const node = (limit: number, bun?: string) => ({
	process: {
		...(bun ? { versions: { bun } } : {}),
		getBuiltinModule: (id: string) =>
			id === 'node:v8' ? { getHeapStatistics: () => ({ heap_size_limit: limit }) } : undefined,
	},
});

describe('runtimeHeapLimit', () => {
	it('reads the V8 heap limit Node and Deno report, and Chromium pages', () => {
		expect(runtimeHeapLimit(node(704 * MiB))).toBe(704 * MiB);
		expect(runtimeHeapLimit({ performance: { memory: { jsHeapSizeLimit: 4096 * MiB } } })).toBe(
			4096 * MiB,
		);
	});
	it('reports nothing where the runtime has no such limit', () => {
		// Bun answers the same call for JavaScriptCore, which has no fixed heap limit.
		expect(runtimeHeapLimit(node(300 * MiB, '1.4.4'))).toBeUndefined();
		expect(runtimeHeapLimit({})).toBeUndefined();
		expect(
			runtimeHeapLimit({
				process: {
					getBuiltinModule: () => {
						throw new Error('no such module');
					},
				},
			}),
		).toBeUndefined();
	});
});

describe('defaultXmlHeapBytes', () => {
	it('leaves the margin out of the heap limit and allows for the peak per byte charged', () => {
		// Node's heap_size_limit under --max-old-space-size=512, 768 and 1024 (young generation added).
		expect(defaultXmlHeapBytes(704 * MiB)).toBe(280 * MiB);
		expect(defaultXmlHeapBytes(960 * MiB)).toBe(440 * MiB);
		expect(defaultXmlHeapBytes(1216 * MiB)).toBe(600 * MiB);
	});
	it('assumes a 1 GiB heap where none is reported, and keeps a floor on a tiny heap', () => {
		expect(defaultXmlHeapBytes(1024 * MiB)).toBe(480 * MiB);
		expect(defaultXmlHeapBytes(200 * MiB)).toBe(8 * MiB);
	});
});
