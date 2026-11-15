import { neutralizeDeviceStrokes } from './device-pixel-stroke';
import { ACTION_INDICATOR_CLASS, LINK_TOOLTIP_CLASS } from './element-action-affordance';

/**
 * Editor affordances every binding paints inside an element's box without marking them
 * `data-export-ignore`: the amber action badge and the hover link tooltip. They are editor
 * chrome (PowerPoint's own exports have neither), so no export draws them.
 */
const EDITOR_AFFORDANCES = `.${ACTION_INDICATOR_CLASS}, .${LINK_TOOLTIP_CLASS}`;

/** Explicit editor-only DOM is omitted by both raster export strategies. */
export function isExportIgnoredElement(element: Element): boolean {
	return (
		element.getAttribute('data-export-ignore') === 'true' ||
		element.matches?.(EDITOR_AFFORDANCES) === true
	);
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
	for (const overlay of root.querySelectorAll(
		`[data-export-ignore="true"], ${EDITOR_AFFORDANCES}`,
	)) {
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
