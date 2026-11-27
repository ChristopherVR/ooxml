// The Draw tab's ribbon icons (24 x 24 line art), registered when the tab loads. A function the tab
// calls rather than an import for its side effect: the package's `sideEffects` list lets bundlers
// drop a bare import, which left the built tab's icons blank.
import { ribbonIconPaths } from './ribbon-icons';

const DRAW_ICONS = {
	lasso: 'M12 4c4.4 0 8 2.2 8 5s-3.6 5-8 5-8-2.2-8-5 3.6-5 8-5z M6 13c-1 2 0 4 2 4 1.5 0 2 1.5 1 3',
	eraser: 'M15 4l5 5-9 9H7l-3-3z M11 8l5 5 M10 20h10',
	addPen: 'M14 4l4 4-9 9H5v-4z M12 6l4 4 M17 15v6 M14 18h6',
	drawTouch:
		'M9 11V5a1.5 1.5 0 0 1 3 0v5 M12 9.5a1.5 1.5 0 0 1 3 0v1 M15 10.5a1.5 1.5 0 0 1 3 0V15c0 3-2 6-5.5 6S8 19 6.5 16L5 13.5a1.4 1.4 0 0 1 2.3-1.6L9 14',
	inkToShape: 'M3 17c3-6 5 2 8-4 M14 6h7v7h-7z M11 13l3-3',
	inkToMath: 'M3 15c2-5 4 3 6-2 M21 5H13l4 5-4 5h8 M10 18h4',
	inkReplay: 'M4 17c3-7 6 3 9-4 M15 7l6 4-6 4z',
};

/** Adds the Draw tab's icons to the ribbon's icon set. Idempotent. */
export function registerDrawIcons(): void {
	Object.assign(ribbonIconPaths, DRAW_ICONS);
}
