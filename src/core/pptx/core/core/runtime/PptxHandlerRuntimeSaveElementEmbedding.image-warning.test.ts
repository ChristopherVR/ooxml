/**
 * `SAVE_IMAGE_PAYLOAD_UNSUPPORTED` must not fire for a loaded image shown through the handler's own
 * blob: URL: its bytes are unchanged and already in the package at `imagePath`. It still fires for
 * a payload the save cannot embed. `processImageEmbedding` is protected deep in the mixin chain, so
 * it is called on a `this`-shaped stub (see PptxHandlerRuntimePresentationProps.test.ts).
 */
import { describe, expect, it, vi } from 'vitest';

import './PptxHandlerRuntimeChartParsingHelpers';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeSaveElementEmbedding';

type Embed = (el: unknown, shape: unknown, ctx: unknown) => unknown;

function embed(imageData: string, imagePath: string | undefined, ownUrls: string[]) {
	const reportWarning = vi.fn();
	const host = {
		blobUrlCache: new Set(ownUrls),
		compatibilityService: { reportWarning },
		parseDataUrlToBytes: () => null,
	};
	const processImageEmbedding = (
		PptxHandlerRuntime.prototype as unknown as { processImageEmbedding: Embed }
	).processImageEmbedding;
	const shape = { 'p:pic': {} };
	const result = processImageEmbedding.call(
		host,
		{ id: 'pic-1', type: 'image', imageData, imagePath },
		shape,
		{ slide: { id: 'ppt/slides/slide1.xml' } },
	);
	return { reportWarning, result, shape };
}

describe('processImageEmbedding warnings', () => {
	it('stays quiet for an unchanged loaded image shown through its own blob URL', () => {
		const { reportWarning, result, shape } = embed(
			'blob:http://localhost/1',
			'ppt/media/image1.png',
			['blob:http://localhost/1'],
		);
		expect(reportWarning).not.toHaveBeenCalled();
		expect(result).toBe(shape);
	});

	it('still warns for a blob URL it did not create', () => {
		const { reportWarning } = embed('blob:http://localhost/other', 'ppt/media/image1.png', []);
		expect(reportWarning).toHaveBeenCalledWith(
			expect.objectContaining({ code: 'SAVE_IMAGE_PAYLOAD_UNSUPPORTED' }),
		);
	});

	it('still warns for its own blob URL when the image has no package part', () => {
		const { reportWarning } = embed('blob:http://localhost/1', undefined, [
			'blob:http://localhost/1',
		]);
		expect(reportWarning).toHaveBeenCalledOnce();
	});
});
