import { afterEach, describe, expect, it, vi } from 'vitest';
import { encodePng } from 'ooxml-core/pptx';
import {
	clearBarFacePicturePixelCache,
	ensureBarFacePicturePixelSampled,
	getCachedBarFacePicturePixelColor,
	resetBarFacePicturePixelCacheForTests,
	resolveBarFacePicturePixelColor,
	subscribeBarFacePicturePixelSamples,
} from './chart-bar3d-face-picture-sample';

async function flush(): Promise<void> {
	for (let i = 0; i < 64; i++) await Promise.resolve();
}
afterEach(() => resetBarFacePicturePixelCacheForTests());

describe('bar-face sample admission', () => {
	it('does not start or notify an oversized source, even when notifications request it again', async () => {
		const source = `blob:${'x'.repeat(4 * 1024 * 1024)}`;
		const sampler = vi.fn(async () => '#123456');
		let notifications = 0;
		const unsubscribe = subscribeBarFacePicturePixelSamples(() => {
			// A broken implementation must fail finitely rather than hang this test.
			if (++notifications < 8) ensureBarFacePicturePixelSampled(source, sampler);
		});
		try {
			ensureBarFacePicturePixelSampled(source, sampler);
			await flush();
			expect(sampler).not.toHaveBeenCalled();
			expect(notifications).toBe(0);
		} finally {
			unsubscribe();
		}
	});

	it('does not repeatedly decode when broadcast-driven renders request 129 active pictures', async () => {
		const sources = Array.from({ length: 129 }, (_, index) => `blob:picture-${index}`);
		const sampler = vi.fn(async () => '#123456');
		let notifications = 0;
		const render = () => {
			for (const source of sources) ensureBarFacePicturePixelSampled(source, sampler);
		};
		const unsubscribe = subscribeBarFacePicturePixelSamples(() => {
			if (++notifications < 512) queueMicrotask(render);
		});
		try {
			render();
			await flush();
			expect(sampler).toHaveBeenCalledTimes(128);
			expect(notifications).toBe(128);
			expect(getCachedBarFacePicturePixelColor(sources[0]!)).toBe('#123456');
			expect(getCachedBarFacePicturePixelColor(sources[128]!)).toBeUndefined();
		} finally {
			unsubscribe();
		}
	});

	it('keeps a new epoch reservation when an old decode settles, and releases admission explicitly', async () => {
		let finishOld!: (color: string) => void, finishCurrent!: (color: string) => void;
		const first = `blob:first-${'x'.repeat(3 * 1024 * 1024)}`;
		const second = `blob:second-${'y'.repeat(3 * 1024 * 1024)}`;
		const third = `blob:third-${'z'.repeat(3 * 1024 * 1024)}`;
		const old = vi.fn(
			() =>
				new Promise<string>((resolve) => {
					finishOld = resolve;
				}),
		);
		const current = vi.fn(
			() =>
				new Promise<string>((resolve) => {
					finishCurrent = resolve;
				}),
		);
		const blocked = vi.fn(async () => '#abcdef');
		ensureBarFacePicturePixelSampled(first, old);
		ensureBarFacePicturePixelSampled(second, current);
		await flush();
		expect(old).toHaveBeenCalledOnce();
		expect(current).not.toHaveBeenCalled();
		clearBarFacePicturePixelCache();
		ensureBarFacePicturePixelSampled(second, current);
		await flush();
		finishOld('#123456');
		await flush();
		ensureBarFacePicturePixelSampled(third, blocked);
		await flush();
		expect(blocked).not.toHaveBeenCalled();
		finishCurrent('#00ff00');
		await flush();
		expect(getCachedBarFacePicturePixelColor(first)).toBeUndefined();
		expect(getCachedBarFacePicturePixelColor(second)).toBe('#00ff00');
		ensureBarFacePicturePixelSampled(third, blocked);
		await flush();
		expect(blocked).not.toHaveBeenCalled();
		clearBarFacePicturePixelCache();
		ensureBarFacePicturePixelSampled(third, blocked);
		await flush();
		expect(blocked).toHaveBeenCalledOnce();
		expect(getCachedBarFacePicturePixelColor(third)).toBe('#abcdef');
	});

	it('still returns synchronously decoded colours after memo admission is full', async () => {
		for (let i = 0; i < 128; i++)
			ensureBarFacePicturePixelSampled(`blob:${i}`, async () => '#123456');
		await flush();
		const source = `data:image/png;base64,${Buffer.from(encodePng(1, 1, new Uint8Array([10, 20, 30, 255]))).toString('base64')}`;
		expect(resolveBarFacePicturePixelColor(source)).toBe('#0a141e');
		expect(getCachedBarFacePicturePixelColor(source)).toBeUndefined();
	});
});
