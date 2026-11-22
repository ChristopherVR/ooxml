import type { PptxSlide } from 'ooxml-core/pptx';
import { SvgExporter } from 'ooxml-core/pptx/converter';
import { renderLocalSvg } from '../../form/local-svg';

export async function renderZoomSlidePreview(
	doc: Document,
	slide: PptxSlide,
	width: number,
	height: number,
	transparent: boolean,
): Promise<string> {
	const ratio = Math.min(1, 360 / Math.max(width, height));
	const svg = SvgExporter.exportSlide(slide, width, height, {
		transparentBackground: transparent,
	});
	const canvas = await renderLocalSvg(
		doc,
		svg,
		Math.max(1, Math.round(width * ratio)),
		Math.max(1, Math.round(height * ratio)),
	);
	return canvas.toDataURL('image/png');
}
