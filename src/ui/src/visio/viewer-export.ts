import type { VisioDocument } from 'ooxml-core/visio';
import { exportPageSvg } from './export-svg';

/** Raster resolution of PNG output and of PDF picture-page fallbacks; large pages are scaled down to the canvas cap. */
export const EXPORT_DPI = 150;
const MAX_CANVAS_SIDE = 4096;

/** Pixel size of a page at `dpi`, scaled down so neither side exceeds the canvas cap. */
export function rasterSize(widthIn: number, heightIn: number, dpi = EXPORT_DPI) {
	const scale = Math.min(dpi, MAX_CANVAS_SIDE / widthIn, MAX_CANVAS_SIDE / heightIn);
	return {
		width: Math.max(1, Math.round(widthIn * scale)),
		height: Math.max(1, Math.round(heightIn * scale)),
	};
}

/** Draw one exported page SVG on a white canvas. Resolves once the image has loaded. */
export async function rasterizePage(
	doc: Document,
	model: VisioDocument,
	pageIndex: number,
	dpi = EXPORT_DPI,
): Promise<HTMLCanvasElement> {
	const exported = exportPageSvg(model, pageIndex);
	const size = rasterSize(exported.width, exported.height, dpi);
	const view = doc.defaultView;
	if (!view) throw new Error('Exporting needs a browser window.');
	const url = view.URL.createObjectURL(
		new Blob([exported.svg], { type: 'image/svg+xml;charset=utf-8' }),
	);
	try {
		const image = new view.Image();
		await new Promise<void>((resolve, reject) => {
			image.onload = () => resolve();
			image.onerror = () => reject(new Error('The page picture could not be drawn.'));
			image.src = url;
		});
		const canvas = doc.createElement('canvas');
		canvas.width = size.width;
		canvas.height = size.height;
		const context = canvas.getContext('2d');
		if (!context) throw new Error('This browser cannot draw the page picture.');
		context.fillStyle = '#ffffff';
		context.fillRect(0, 0, size.width, size.height);
		context.drawImage(image, 0, 0, size.width, size.height);
		return canvas;
	} finally {
		view.URL.revokeObjectURL(url);
	}
}

const blob = (canvas: HTMLCanvasElement, type: string, quality?: number) =>
	new Promise<Blob>((resolve, reject) =>
		canvas.toBlob(
			(result) =>
				result ? resolve(result) : reject(new Error('The picture could not be encoded.')),
			type,
			quality,
		),
	);

/** The current page as a PNG picture. */
export async function exportPagePng(
	doc: Document,
	model: VisioDocument,
	pageIndex: number,
): Promise<Blob> {
	return blob(await rasterizePage(doc, model, pageIndex), 'image/png');
}
