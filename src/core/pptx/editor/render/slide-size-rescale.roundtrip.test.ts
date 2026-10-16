import { readFile } from 'node:fs/promises';

import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import { scaleSlidesForSizeChange } from './slide-size-rescale';

const FOUR_THREE = { widthEmu: 9144000, heightEmu: 6858000 };
const WIDESCREEN = { widthEmu: 12192000, heightEmu: 6858000, type: '' };
const fixture = new URL('../../__tests__/fixtures/slide-size-centering.pptx', import.meta.url);

describe('slide-size content centering round trip', () => {
	it.each([
		{ mode: 'ensureFit' as const, frame: { x: 340, y: 210, width: 600, height: 240 } },
		{ mode: 'maximize' as const, frame: { x: 240, y: 160, width: 800, height: 320 } },
	])('keeps content on the resized slide after $mode and save/reload', async ({ mode, frame }) => {
		const handler = new PptxHandler();
		const reloader = new PptxHandler();
		try {
			const bytes = await readFile(fixture);
			const data = await handler.load(Uint8Array.from(bytes).buffer);
			expect(data.width).toBe(960);
			expect(data.height).toBe(720);
			expect(data.slides[0].elements[0]).toMatchObject({ x: 180, y: 210 });

			const resized = scaleSlidesForSizeChange(data.slides, FOUR_THREE, WIDESCREEN, mode);
			const saved = await handler.save(resized, { slideSize: WIDESCREEN });
			const reloaded = await reloader.load(Uint8Array.from(saved).buffer);
			expect(reloaded.width).toBe(1280);
			expect(reloaded.height).toBe(720);
			const shape = reloaded.slides[0].elements[0];
			for (const key of ['x', 'y', 'width', 'height'] as const) {
				expect(shape[key], key).toBeCloseTo(frame[key], 3);
			}
			expect(shape.x).toBeGreaterThanOrEqual(0);
			expect(shape.y).toBeGreaterThanOrEqual(0);
			expect(shape.x + shape.width).toBeLessThanOrEqual(reloaded.width);
			expect(shape.y + shape.height).toBeLessThanOrEqual(reloaded.height);
			expect(data.slides[0].elements[0]).toMatchObject({ x: 180, y: 210 });
		} finally {
			handler.dispose();
			reloader.dispose();
		}
	});
});
