/**
 * Visio's F6 ring: the ribbon, the Shapes window, the drawing window, the task pane, the page tabs
 * and the status bar, in the order Microsoft Visio moves round the window. The walk itself is the
 * shared region service, so every product moves the same way.
 */
import { moveRegionFocus, type WindowRegion } from '../region-focus';

const VISIO_REGIONS: readonly WindowRegion[] = [
	{ selector: 'office-ui-ribbon' },
	{ selector: '#shapes-pane' },
	// The drawing window takes the keyboard itself; Tab then walks the shapes on the page.
	{
		selector: '.viewport',
		focus: (canvas) => {
			canvas.focus({ preventScroll: true });
			return true;
		},
	},
	{ selector: '#inspector-pane' },
	// Visio's bottom chrome: the page tabs with the All pages list, then the status bar.
	{ selector: '.page-bar' },
	{ selector: 'office-ui-status-bar' },
];

/** Move the keyboard to the next (`1`) or previous (`-1`) region of the drawing window. */
export function moveVisioRegion(root: ShadowRoot, direction: 1 | -1): boolean {
	return moveRegionFocus(root, VISIO_REGIONS, direction);
}
