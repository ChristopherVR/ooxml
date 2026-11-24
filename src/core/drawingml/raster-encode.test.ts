import { describe, expect, it } from 'vitest';
import { decodePng } from './raster-decode';
import { encodePng } from './raster-encode';

describe('encodePng', () => {
	it('writes RGB and RGBA pictures that decode to the same samples', () => {
		const rgb = Uint8Array.from([
			255, 0, 0, 0, 255, 0, 0, 0, 255, 9, 9, 9, 10, 20, 30, 200, 100, 50,
		]);
		const opaque = decodePng(encodePng(3, 2, rgb))!;
		expect([...opaque.rgb]).toEqual([...rgb]);
		expect(opaque.alpha).toBeUndefined();
		const alpha = Uint8Array.from([255, 0, 128, 255, 255, 1]);
		const translucent = decodePng(encodePng(3, 2, rgb, alpha))!;
		expect([...translucent.rgb]).toEqual([...rgb]);
		expect([...translucent.alpha!]).toEqual([...alpha]);
	});
});
