import { describe, expect, it, vi } from 'vitest';

import { convertMetafileToDataUrl } from './metafile-convert';

vi.mock('emf-converter', () => ({
	convertMetafileToDataUrl: vi.fn(async () => 'data:image/png;base64,AA=='),
}));

describe('convertMetafileToDataUrl', () => {
	it('loads emf-converter on first call and forwards every argument', async () => {
		const emf = await import('emf-converter');
		const buffer = new ArrayBuffer(4);

		await expect(convertMetafileToDataUrl(buffer, undefined, 1)).resolves.toBe(
			'data:image/png;base64,AA==',
		);
		expect(emf.convertMetafileToDataUrl).toHaveBeenCalledWith(buffer, undefined, 1);
	});
});
