import { expect, it } from 'vitest';
import { PptxHandler } from '../../PptxHandler';
import { createPhotoAlbum } from './photo-album';

const png =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jHuoAAAAASUVORK5CYII=';

it('writes a standalone deck with a title, grouped photos and editable captions', async () => {
	const bytes = await createPhotoAlbum({
		title: 'Field trip',
		picturesPerSlide: 2,
		captions: true,
		pictures: ['Terrace', 'Waterfall', 'Lookout'].map((caption) => ({
			data: png,
			width: 1200,
			height: 800,
			caption,
		})),
	});
	const handler = new PptxHandler();
	try {
		const deck = await handler.load(bytes.slice().buffer, { eagerDecodeImages: true });
		expect(deck.slides).toHaveLength(3);
		expect(deck.slides[0]?.elements).toEqual([expect.objectContaining({ text: 'Field trip' })]);
		const photos = deck.slides[1]?.elements.filter(
			(element) => element.type === 'picture' || element.type === 'image',
		);
		expect(photos).toEqual([
			expect.objectContaining({
				altText: 'Terrace',
				imagePath: expect.stringMatching(/^ppt\/media\/image\d+\.png$/u),
			}),
			expect.objectContaining({ altText: 'Waterfall' }),
		]);
		expect(deck.slides[1]?.elements).toEqual(
			expect.arrayContaining([expect.objectContaining({ text: 'Waterfall' })]),
		);
		expect(deck.slides[2]?.elements).toEqual(
			expect.arrayContaining([expect.objectContaining({ text: 'Lookout' })]),
		);
	} finally {
		handler.dispose();
	}
});
