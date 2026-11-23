import type { ImagePptxElement } from 'ooxml-core/pptx';
import { renderLocalSvg } from '../../form/local-svg';
import { generateElementId } from '../render/element-clipboard';
import type { CanvasSize } from '../types';

const POSTER =
	'<svg xmlns="http://www.w3.org/2000/svg" width="192" height="192" viewBox="0 0 192 192"><rect width="192" height="192" fill="#28394c"/><circle cx="96" cy="72" r="28" fill="#9eafc3"/><path d="M40 166v-10a56 56 0 0 1 112 0v10Z" fill="#9eafc3"/></svg>';

export async function createCameoElement(
	doc: Document,
	{ width, height }: CanvasSize,
	name: string,
): Promise<ImagePptxElement> {
	const size = Math.min(width, height) * 0.3;
	const poster = await renderLocalSvg(doc, POSTER, 192, 192);
	return {
		type: 'image',
		id: generateElementId(),
		name,
		altText: name,
		x: width - size - width * 0.05,
		y: height - size - height * 0.05,
		width: size,
		height: size,
		shapeType: 'ellipse',
		cropShape: 'ellipse',
		imageData: poster.toDataURL('image/png'),
		liveFeed: true,
	};
}
