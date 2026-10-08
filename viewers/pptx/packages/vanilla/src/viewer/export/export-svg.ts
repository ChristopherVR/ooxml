import type { PptxData, PptxSlide, SvgExportOptions as CoreSvgExportOptions } from 'pptx-viewer-core';
import { SvgExporter } from 'pptx-viewer-core';

/**
 * Options for the SVG exporters. A local alias rather than a re-export of
 * `pptx-viewer-core`: the declaration bundler wrote a re-export reached
 * through `export *` as a dotted name in an export list, which is a syntax
 * error for consumers (issue #33).
 */
export type SvgExportOptions = CoreSvgExportOptions;

/** Export one parsed slide as resolution-independent SVG markup. */
export function exportSlideToSvg(
	slide: PptxSlide,
	width: number,
	height: number,
	options: SvgExportOptions = {},
): string {
	return SvgExporter.exportSlide(slide, width, height, options);
}

/** Export the selected slides in a parsed presentation as SVG markup. */
export function exportAllSlidesToSvg(data: PptxData, options: SvgExportOptions = {}): string[] {
	return SvgExporter.exportAll(data, options);
}

