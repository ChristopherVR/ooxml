/**
 * Node.js SVG -> PNG rasteriser for the `.ppt` writer, through the optional
 * `@napi-rs/canvas` peer (the same Node canvas backend `emf-converter`
 * already falls back to), whose `loadImage` decodes SVG with Skia's SVG
 * module. Returns `undefined` when the package is not installed, so the
 * writer keeps degrading to a placeholder with a warning.
 *
 * Only the Node-only `ooxml-core/pptx/node` entry reaches this module: the
 * engine never imports it, so browser bundles never see the package.
 *
 * @module pptx/node/svg-rasterize-node
 */

import type { PptxSvgRasterizer } from '../core/core/types';

/** The slice of `@napi-rs/canvas` this module uses. */
interface NodeCanvasImage {
	readonly width: number;
	readonly height: number;
}
interface NodeCanvasContext {
	drawImage(image: NodeCanvasImage, dx: number, dy: number, dw: number, dh: number): void;
}
interface NodeCanvas {
	getContext(kind: '2d'): NodeCanvasContext;
	toBuffer(mime: 'image/png'): Uint8Array;
}
interface NodeCanvasModule {
	createCanvas(width: number, height: number): NodeCanvas;
	loadImage(source: Uint8Array): Promise<NodeCanvasImage>;
}

let nodeCanvas: Promise<NodeCanvasModule | undefined> | undefined;

function isNodeCanvasModule(value: unknown): value is NodeCanvasModule {
	const candidate = value as Partial<NodeCanvasModule> | null;
	return typeof candidate?.createCanvas === 'function' && typeof candidate.loadImage === 'function';
}

function loadNodeCanvas(): Promise<NodeCanvasModule | undefined> {
	nodeCanvas ??= (async () => {
		try {
			const imported: unknown = await import('@napi-rs/canvas');
			const withDefault = imported as { default?: unknown };
			if (isNodeCanvasModule(imported)) {
				return imported;
			}
			return isNodeCanvasModule(withDefault.default) ? withDefault.default : undefined;
		} catch {
			return undefined;
		}
	})();
	return nodeCanvas;
}

/**
 * Rasterise SVG bytes to PNG through `@napi-rs/canvas`, sized by `sizeFor`
 * from the SVG's intrinsic size (0 when it declares none). Pass it as
 * `rasterizeSvg` to a `.ppt` save.
 *
 * @returns PNG bytes, or `undefined` without the package or on a bad SVG.
 */
export const rasterizeSvgInNode: PptxSvgRasterizer = async (svg, sizeFor) => {
	const canvasModule = await loadNodeCanvas();
	if (!canvasModule) {
		return undefined;
	}
	try {
		const image = await canvasModule.loadImage(svg);
		const { width, height } = sizeFor(image.width, image.height);
		const canvas = canvasModule.createCanvas(width, height);
		canvas.getContext('2d').drawImage(image, 0, 0, width, height);
		return new Uint8Array(canvas.toBuffer('image/png'));
	} catch {
		return undefined;
	}
};
