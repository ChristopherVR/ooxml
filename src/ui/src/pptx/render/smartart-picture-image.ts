import type { PptxSmartArtDrawingShape } from 'ooxml-core/pptx';
import { pictureSvg } from '../../picture-view';
import { getPresetShapeVectorGeometry } from './stroke-outline';
import { cachedPicturePixels, needsPicturePixels, probePicturePixels } from '../../picture-pixels';
import { getCachedNativeImageSize, probeNativeImageSize } from '../../image-native-size';

export const needsSmartArtPictureRefresh = (shape: PptxSmartArtDrawingShape): boolean =>
	Boolean(
		shape.fillImageUrl &&
		(needsPicturePixels(shape.pictureFill?.effects ?? []) ||
			(shape.pictureFill?.layout?.tile && !(shape.fillImageWidth && shape.fillImageHeight))),
	);

const pendingImages = new WeakMap<SVGImageElement, { key: string; href: string | null }>();
const sourceKey = (shape: PptxSmartArtDrawingShape) =>
	JSON.stringify([
		shape.fillImageUrl,
		shape.pictureFill,
		shape.width,
		shape.height,
		shape.rotation,
		shape.flipHorizontal,
		shape.flipVertical,
		shape.shapeType,
		shape.shapeAdjustments,
	]);

/** Native bindings invalidate only their mounted image once asynchronous pixel paint is ready. */
export function refreshSmartArtPictureImage(
	event: { currentTarget: EventTarget | null },
	source?: PptxSmartArtDrawingShape,
): void {
	const image = event.currentTarget as SVGImageElement | null;
	if (!image || image.localName !== 'image' || !source || !needsSmartArtPictureRefresh(source))
		return;
	const key = sourceKey(source);
	const originalHref = image.getAttribute('href');
	const previous = pendingImages.get(image);
	if (previous?.key === key && previous.href === originalHref) return;
	const state = { key, href: originalHref };
	pendingImages.set(image, state);
	const pixels = needsPicturePixels(source.pictureFill?.effects ?? []);
	const ready = pixels
		? probePicturePixels(
				image.ownerDocument,
				source.fillImageUrl!,
				source.pictureFill!.effects ?? [],
			)
		: probeNativeImageSize(source.fillImageUrl!);
	void ready.then((result) => {
		if (
			!result ||
			!image.isConnected ||
			image.getAttribute('href') !== originalHref ||
			pendingImages.get(image) !== state ||
			sourceKey(source) !== key
		)
			return;
		image.setAttribute('data-office-picture-paint-ready', 'true');
		if (pixels) image.setAttribute('data-office-picture-pixels-ready', 'true');
		const url = smartArtPictureImage(source);
		state.href = url ?? originalHref;
		if (url && image.getAttribute('href') !== url) image.setAttribute('href', url);
	});
}

/** Self-contained image descriptors are shared by every native framework template. */
export function smartArtPictureImage(shape: PptxSmartArtDrawingShape): string | undefined {
	const url = shape.fillImageUrl;
	if (!url?.startsWith('data:') || !shape.pictureFill || typeof document === 'undefined')
		return url;
	const size =
		cachedPicturePixels(url, shape.pictureFill.effects ?? []) ?? getCachedNativeImageSize(url);
	const width = shape.fillImageWidth ?? size?.width;
	const height = shape.fillImageHeight ?? size?.height;
	// Unknown raster formats retain the existing image fallback. Do not serialize an empty
	// asynchronously populated tile pattern into a detached SVG image descriptor.
	if (shape.pictureFill.layout?.tile && !(width && height)) return url;
	const svg = pictureSvg(
		document,
		shape.pictureFill,
		{
			url,
			...(width ? { width } : {}),
			...(height ? { height } : {}),
		},
		shape.width,
		shape.height,
		{
			rotation: shape.rotation,
			flipHorizontal: shape.flipHorizontal,
			flipVertical: shape.flipVertical,
		},
	);
	svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
	const geometry = getPresetShapeVectorGeometry(
		shape.shapeType,
		shape.width,
		shape.height,
		shape.shapeAdjustments,
	);
	const rectangle = svg.lastElementChild;
	if (geometry && rectangle) {
		const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
		path.setAttribute('d', geometry.d);
		path.setAttribute('fill', rectangle.getAttribute('fill') ?? 'none');
		rectangle.replaceWith(path);
	}
	return `data:image/svg+xml,${encodeURIComponent(svg.outerHTML)}`;
}
