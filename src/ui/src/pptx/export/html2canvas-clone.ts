/**
 * The html2canvas-pro `onclone` pass every binding runs before capture.
 *
 * Each binding keeps only the thin `renderToCanvas` glue that imports
 * `html2canvas-pro`; everything it does to the cloned document lives here so
 * the five bindings cannot drift. They did: React and Angular dropped
 * editor-only nodes (`data-export-ignore`: collaboration cursors, remote
 * selections, AI change ghosts) from the clone while Vue, Svelte and Vanilla
 * captured them.
 *
 * A `<pptx-three-view>` is swapped for its pixels here too, the same step the
 * `foreignObject` path takes. html2canvas-pro does clone an open shadow root
 * (and copies its canvas), so 3D views already exported through it; the swap
 * makes that independent of the library's shadow-DOM cloning and settles any
 * view still loading before capture (GIF/video frames included).
 *
 * Pure DOM work with no html2canvas import, so it is unit-testable in jsdom.
 *
 * @module export/html2canvas-clone
 */
import { prepareExportClone } from '../render/export-clone';
import { settleThreeViews, snapshotThreeViewsIntoClone } from '../three-view/export-snapshot';
import { normalizeColorsForCapture } from './canvas-color-fix';
import { preprocessCssForCapture } from './css-preprocessing';

/**
 * Prepare html2canvas's cloned capture document.
 *
 * `original` is the live element handed to html2canvas; `doc` and `clonedEl`
 * are what its `onclone` callback receives. Order matters:
 *
 * 0. The cloned element keeps the live element's layout size (see
 *    `html2canvasCaptureScope`).
 * 1. 3D views settle on the live tree, then each cloned view is replaced by an
 *    `<img>` of its canvas plus its DOM overlay. This runs first, while the
 *    live and cloned trees still pair up in document order.
 * 2. Editor-only nodes (`data-export-ignore`) are removed and the recorded
 *    pre-selection paint is restored, exactly as the `foreignObject` path does.
 * 3. Colours html2canvas cannot parse are normalised to sRGB.
 * 4. CSS features html2canvas cannot paint are flattened.
 */
export async function prepareHtml2CanvasClone(
	original: HTMLElement,
	doc: Document,
	clonedEl: HTMLElement,
): Promise<void> {
	await settleThreeViews(original);
	pinCapturedSize(original, clonedEl);
	snapshotThreeViewsIntoClone(original, clonedEl);
	prepareExportClone(clonedEl);
	await normalizeColorsForCapture(doc, clonedEl);
	preprocessCssForCapture(clonedEl);
}

/** One step up the flattened tree html2canvas clones: a slotted node's slot, a parent, or a host. */
function composedParent(node: Node): Node | null {
	const slot = (node as Element | Text).assignedSlot;
	if (slot) return slot;
	const parent = node.parentNode;
	return parent?.nodeType === 11 && (parent as ShadowRoot).host
		? (parent as ShadowRoot).host
		: parent;
}

const STYLE_SOURCES = 'style, link[rel~="stylesheet"]';

/**
 * The `ignoreElements` every binding hands html2canvas-pro.
 *
 * html2canvas-pro clones the whole document, every element and open shadow root, before it paints
 * the one element it captures. Under a viewer with its chrome mounted that clone is most of a
 * capture (seconds, all on the main thread), and nothing outside the captured element is painted.
 * This keeps the captured element's subtree (through shadow roots and slots), the chain of its
 * ancestors, whose styles it inherits, and every style sheet with the path to it (the clone lays
 * ordinary elements out again from those rules, fonts included); it skips every other branch. The
 * clone is laid out without those neighbours, so `prepareHtml2CanvasClone` keeps the captured
 * element at its live size. `ignore`, a caller's own predicate, still applies.
 */
export function html2canvasCaptureScope(
	captured: Element,
	ignore?: (node: Element) => boolean,
): (node: Element) => boolean {
	const ancestors = new Set<Node>();
	for (let node = composedParent(captured); node; node = composedParent(node)) {
		ancestors.add(node);
	}
	const inside = (node: Node): boolean => {
		for (let current: Node | null = node; current; current = composedParent(current)) {
			if (current === captured) return true;
			if (ancestors.has(current)) return false;
		}
		return false;
	};
	return (node) => {
		if (ignore?.(node)) return true;
		if (ancestors.has(node) || inside(node)) return false;
		return !(
			node.closest('head') ||
			node.matches(STYLE_SOURCES) ||
			node.querySelector(STYLE_SOURCES)
		);
	};
}

/**
 * Keep the cloned capture at the live element's layout size. Scoped by `html2canvasCaptureScope`,
 * its ancestors are laid out without their other children, which could otherwise resize it (a grid
 * column, a flex share). The offset size is the layout box, before any transform.
 */
export function pinCapturedSize(original: HTMLElement, clonedEl: HTMLElement): void {
	const { offsetWidth: width, offsetHeight: height } = original;
	if (width <= 0 || height <= 0) return;
	const display = original.ownerDocument.defaultView?.getComputedStyle(original).display;
	if (display === 'inline' || display === 'contents') return;
	Object.assign(clonedEl.style, {
		boxSizing: 'border-box',
		width: `${width}px`,
		height: `${height}px`,
		minWidth: '0',
		minHeight: '0',
		maxWidth: 'none',
		maxHeight: 'none',
		flex: 'none',
	});
}
