import { neutralizeDeviceStrokes } from './device-pixel-stroke';

/** Explicit editor-only DOM is omitted by both raster export strategies. */
export function isExportIgnoredElement(element: Element): boolean {
	return element.getAttribute('data-export-ignore') === 'true';
}

/**
 * Prepare a detached capture clone, never the live slide. Inline selection
 * decoration can share the authored element's node, so renderers record its
 * original paint instead of asking export to guess which shadows are editor UI.
 *
 * On-screen strokes are widened to one device pixel of the LIVE zoom (see
 * `device-pixel-stroke`); an export renders at its own resolution and keeps the
 * authored widths, so the clone pins that device pixel to `0px`.
 */
export function prepareExportClone(root: HTMLElement): void {
	neutralizeDeviceStrokes(root);
	for (const overlay of root.querySelectorAll('[data-export-ignore="true"]')) {
		overlay.remove();
	}
	for (const element of [root, ...root.querySelectorAll<HTMLElement | SVGElement>('*')]) {
		for (const property of ['outline', 'outline-offset', 'box-shadow']) {
			const original = element.getAttribute(`data-export-original-${property}`);
			if (original !== null) {
				element.style.setProperty(property, original);
			}
		}
	}
}
