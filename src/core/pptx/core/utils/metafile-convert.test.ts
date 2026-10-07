import { describe, expect, it, vi } from 'vitest';

const converter = vi.hoisted(() => ({
	loaded: false,
	convert: vi.fn(async () => 'data:image/png;base64,AA=='),
}));

vi.mock('emf-converter', () => {
	converter.loaded = true;
	return { convertMetafileToDataUrl: converter.convert };
});

describe('convertMetafileToDataUrl', () => {
	it('loads emf-converter only on the first call and forwards every argument', async () => {
		const { convertMetafileToDataUrl } = await import('./metafile-convert');
		expect(converter.loaded).toBe(false);

		const buffer = new ArrayBuffer(4);
		await expect(convertMetafileToDataUrl(buffer, undefined, 1)).resolves.toBe(
			'data:image/png;base64,AA==',
		);
		expect(converter.loaded).toBe(true);
		expect(converter.convert).toHaveBeenCalledWith(buffer, undefined, 1);
	});
});
