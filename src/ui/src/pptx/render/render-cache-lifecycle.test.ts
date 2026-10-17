import { afterEach, describe, expect, it, vi } from 'vitest';
import { getCachedResult, setCachedResult } from './image-color-change';
import { getDuotoneCachedResult, setDuotoneCachedResult } from './image-duotone-canvas';
import {
	_resetNativeImageSizeCacheForTests,
	getCachedNativeImageSize,
	probeNativeImageSize,
} from './image-native-size';
import {
	ensureBarFacePicturePixelSampled,
	getCachedBarFacePicturePixelColor,
	resetBarFacePicturePixelCacheForTests,
} from './chart-bar3d-face-picture-sample';

class LoadedImage {
	naturalWidth = 4;
	naturalHeight = 2;
	onload: (() => void) | null = null;
	set src(_value: string) {
		queueMicrotask(() => this.onload?.());
	}
}

afterEach(() => {
	releasePptxRenderCaches();
	_resetNativeImageSizeCacheForTests();
	resetBarFacePicturePixelCacheForTests();
	vi.unstubAllGlobals();
});

describe('render cache bounds', () => {
	it('evicts old intrinsic dimensions after 128 different sources', async () => {
		vi.stubGlobal('Image', LoadedImage);
		for (let i = 0; i < 129; i++) await probeNativeImageSize(`blob:${i}`);
		expect(getCachedNativeImageSize('blob:0')).toBeUndefined();
		expect(getCachedNativeImageSize('blob:128')).toEqual({ width: 4, height: 2 });
	});

	it('evicts old sampled picture colours after 128 different sources', async () => {
		for (let i = 0; i < 129; i++) {
			ensureBarFacePicturePixelSampled(`blob:${i}`, async () => '#123456');
			await flush();
			expect(getCachedBarFacePicturePixelColor(`blob:${i}`)).toBe('#123456');
		}
		expect(getCachedBarFacePicturePixelColor('blob:0')).toBeUndefined();
	});

	it.each([
		['colour change', setCachedResult, getCachedResult],
		['duotone', setDuotoneCachedResult, getDuotoneCachedResult],
	] as const)('bounds %s image results by bytes as well as entry count', (_name, set, get) => {
		const first = `${_name}-large-1`,
			second = `${_name}-large-2`;
		const dataUrl = 'x'.repeat(3 * 1024 * 1024);
		set(first, dataUrl);
		set(second, dataUrl);
		expect(get(first)).toBeUndefined();
		expect(get(second)).toBe(dataUrl);
		set('oversized', 'x'.repeat(5 * 1024 * 1024));
		expect(get('oversized')).toBeUndefined();
	});
});

import { releasePptxRenderCaches } from './render-cache-lifecycle';
import {
	getBarFacePicturePixelSampleVersion,
	subscribeBarFacePicturePixelSamples,
} from './chart-bar3d-face-picture-sample';
import { resolveColorChangedImageSource } from './image-source-effects';
import { buildCacheKey, DEFAULT_COLOR_CHANGE_TOLERANCE } from './image-color-change';

function deferred<T>() {
	let resolve!: (value: T) => void;
	const promise = new Promise<T>((done) => {
		resolve = done;
	});
	return { promise, resolve };
}
async function flush(): Promise<void> {
	for (let i = 0; i < 6; i++) await Promise.resolve();
}

describe('releasePptxRenderCaches', () => {
	it('does not let an old image decode repopulate or remove a newer pending probe', async () => {
		const images: LoadedImage[] = [];
		class PendingImage extends LoadedImage {
			constructor() {
				super();
				images.push(this);
			}
			set src(_value: string) {}
		}
		vi.stubGlobal('Image', PendingImage);
		const old = probeNativeImageSize('blob:shared');
		releasePptxRenderCaches();
		const current = probeNativeImageSize('blob:shared');
		images[0]!.onload?.();
		await old;
		expect(getCachedNativeImageSize('blob:shared')).toBeUndefined();
		expect(probeNativeImageSize('blob:shared')).toBe(current);
		images[1]!.onload?.();
		await current;
		expect(getCachedNativeImageSize('blob:shared')).toEqual({ width: 4, height: 2 });
	});

	it('suppresses late samples while preserving subscriptions and newer requests', async () => {
		const old = deferred<string>(),
			current = deferred<string>();
		let notifications = 0;
		const unsubscribe = subscribeBarFacePicturePixelSamples(() => {
			notifications++;
		});
		try {
			ensureBarFacePicturePixelSampled('blob:shared', () => old.promise);
			releasePptxRenderCaches();
			const version = getBarFacePicturePixelSampleVersion();
			ensureBarFacePicturePixelSampled('blob:shared', () => current.promise);
			old.resolve('#ff0000');
			await flush();
			expect(getCachedBarFacePicturePixelColor('blob:shared')).toBeUndefined();
			expect(getBarFacePicturePixelSampleVersion()).toBe(version);
			const duplicate = vi.fn();
			ensureBarFacePicturePixelSampled('blob:shared', duplicate);
			expect(duplicate).not.toHaveBeenCalled();
			current.resolve('#00ff00');
			await flush();
			expect(getCachedBarFacePicturePixelColor('blob:shared')).toBe('#00ff00');
			expect(notifications).toBe(1);
		} finally {
			unsubscribe();
		}
	});

	it.each([
		['colour change', setCachedResult, getCachedResult],
		['duotone', setDuotoneCachedResult, getDuotoneCachedResult],
	] as const)('clears %s results while allowing new explicit cache writes', (_name, set, get) => {
		set('existing', 'data:old');
		releasePptxRenderCaches();
		expect(get('existing')).toBeUndefined();
		set('fresh', 'data:fresh');
		expect(get('fresh')).toBe('data:fresh');
	});

	it('returns an in-flight colour effect to its caller without retaining it after release', async () => {
		let image!: LoadedImage;
		class PendingImage extends LoadedImage {
			constructor() {
				super();
				image = this;
			}
			set src(_value: string) {}
		}
		vi.stubGlobal('Image', PendingImage);
		vi.stubGlobal('document', {
			createElement: () => ({
				getContext: () => ({
					drawImage() {},
					putImageData() {},
					getImageData: () => ({ data: new Uint8ClampedArray([0, 0, 0, 255]) }),
				}),
				toDataURL: () => 'data:processed',
			}),
		});
		const source = 'blob:effect';
		const pending = resolveColorChangedImageSource(source, {
			clrFrom: '#000000',
			clrTo: '#ffffff',
		});
		releasePptxRenderCaches();
		image.onload?.();
		await expect(pending).resolves.toBe('data:processed');
		const key = buildCacheKey(source, '#000000', '#ffffff', DEFAULT_COLOR_CHANGE_TOLERANCE, false);
		expect(getCachedResult(key)).toBeUndefined();
	});
});
