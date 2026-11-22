import type { DisplayRegion } from './display-clipping';
import type { previewDisplay } from './display-capture';

export function cropDisplayStream(
	preview: Awaited<ReturnType<typeof previewDisplay>>,
	source: MediaStream,
	region: DisplayRegion,
) {
	const { image, video } = preview;
	const canvas = image.ownerDocument.createElement('canvas');
	canvas.width = region.width;
	canvas.height = region.height;
	const context = canvas.getContext('2d');
	if (!context) throw new Error('Could not render the selected screen area');

	const draw = () => {
		if (video.readyState < HTMLMediaElement.HAVE_CURRENT_DATA) return;
		const scaleX = video.videoWidth / image.width;
		const scaleY = video.videoHeight / image.height;
		context.drawImage(
			video,
			region.x * scaleX,
			region.y * scaleY,
			region.width * scaleX,
			region.height * scaleY,
			0,
			0,
			canvas.width,
			canvas.height,
		);
	};
	draw();
	const stream = canvas.captureStream(30);
	for (const track of source.getAudioTracks()) stream.addTrack(track);
	// A shared window or tab can put the editor in the background.
	const timer = setInterval(draw, 1000 / 30);
	return { stream, release: () => clearInterval(timer) };
}
