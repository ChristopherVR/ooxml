import type { PptxSmartArtDrawingShape } from 'ooxml-core/pptx';
import { pictureSvg } from '../../picture-view';
import { getPresetShapeVectorGeometry } from './stroke-outline';

/** Self-contained image descriptors are shared by every native framework template. */
export function smartArtPictureImage(shape: PptxSmartArtDrawingShape): string | undefined {
	const url = shape.fillImageUrl;
	if (!url?.startsWith('data:') || !shape.pictureFill || typeof document === 'undefined')
		return url;
	// Unknown raster formats retain the existing image fallback. Do not serialize an empty
	// asynchronously populated tile pattern into a detached SVG image descriptor.
	if (shape.pictureFill.layout?.tile && !(shape.fillImageWidth && shape.fillImageHeight))
		return url;
	const svg = pictureSvg(
		document,
		shape.pictureFill,
		{
			url,
			...(shape.fillImageWidth ? { width: shape.fillImageWidth } : {}),
			...(shape.fillImageHeight ? { height: shape.fillImageHeight } : {}),
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
