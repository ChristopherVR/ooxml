/**
 * The size an exported element is rasterised at, and whether its own CSS transform is only a zoom.
 *
 * A slide stage on screen is laid out at the slide's size and zoomed to fit with a uniform
 * `scale` transform. Sizing an export from its on-screen box made the exported resolution follow
 * the window and the zoom, and drew the zoomed stage into the corner of the canvas. A stage is
 * therefore exported at its layout size with that transform dropped (`untransformedRoot`). Any
 * other transform (a rotated or flipped shape exported as a picture) is part of the picture, so
 * such an element keeps its on-screen box and its transform.
 *
 * @module export/export-root-size
 */

/** What {@link measureExportRoot} decided for one element. */
export interface ExportRootSize {
	/** CSS px width to rasterise. */
	width: number;
	/** CSS px height to rasterise. */
	height: number;
	/** Draw the element without its own transform: it is only a zoom. */
	untransformedRoot: boolean;
}

/** The scale of a computed `transform` that only zooms (uniform, positive, unrotated), or undefined. */
export function uniformZoomOf(transform: string): number | undefined {
	const match = /^matrix\(([^)]*)\)$/u.exec(transform.trim());
	if (!match) {
		return undefined;
	}
	const [a, b, c, d] = (match[1] ?? '').split(',').map(Number);
	if (a === undefined || b === undefined || c === undefined || d === undefined) {
		return undefined;
	}
	return a > 0 && b === 0 && c === 0 && Math.abs(a - d) < 1e-9 ? a : undefined;
}

/** See the module comment. */
export function measureExportRoot(element: HTMLElement): ExportRootSize {
	const rect = element.getBoundingClientRect();
	const transform =
		element.ownerDocument.defaultView?.getComputedStyle(element).transform ?? 'none';
	if (
		uniformZoomOf(transform) !== undefined &&
		element.offsetWidth > 0 &&
		element.offsetHeight > 0
	) {
		return { width: element.offsetWidth, height: element.offsetHeight, untransformedRoot: true };
	}
	return {
		width: rect.width || element.offsetWidth,
		height: rect.height || element.offsetHeight,
		untransformedRoot: false,
	};
}

/** For an `html2canvas` `onclone`: drop the cloned root's zoom, as `untransformedRoot` does. */
export function untransformExportClone(clone: HTMLElement): void {
	for (const property of ['transform', 'translate', 'rotate', 'scale']) {
		clone.style.setProperty(property, 'none');
	}
}
