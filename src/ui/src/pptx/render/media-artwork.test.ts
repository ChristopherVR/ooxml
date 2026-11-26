import type { MediaPptxElement } from 'ooxml-core/pptx';
import { expect, it } from 'vitest';

import { mediaArtworkClipPath } from './media-artwork';

const video: MediaPptxElement = {
	id: 'clip',
	type: 'media',
	mediaType: 'video',
	x: 0,
	y: 0,
	width: 160,
	height: 90,
};

it('clips media artwork to the preset silhouette outside fullscreen playback', () => {
	expect(mediaArtworkClipPath({ ...video, shapeType: 'rect' })).toBeUndefined();
	expect(mediaArtworkClipPath({ ...video, shapeType: 'ellipse' })).toBe('ellipse(50% 50%)');
	expect(mediaArtworkClipPath({ ...video, shapeType: 'hexagon' })).toBeTruthy();
	expect(mediaArtworkClipPath({ ...video, shapeType: 'ellipse' }, true)).toBeUndefined();
});
