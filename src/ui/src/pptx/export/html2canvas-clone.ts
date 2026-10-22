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
 * 0. The cloned element keeps the live element's layout width (see
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

/** Whether a style sheet sits in `node` or in a node one of its slots shows. */
function holdsStyleSource(node: Element): boolean {
	if (node.matches(STYLE_SOURCES) || node.querySelector(STYLE_SOURCES)) return true;
	return [...node.querySelectorAll('slot')].some((slot) =>
		slot.assignedElements().some((assigned) => holdsStyleSource(assigned)),
	);
}

/**
 * The `ignoreElements` every binding hands html2canvas-pro.
 *
 * html2canvas-pro clones the whole document, every element and open shadow root, before it paints
 * the one element it captures. Under a viewer with its chrome mounted that clone is most of a
 * capture (seconds, all on the main thread), and nothing outside the captured element is painted.
 * This keeps the captured element's subtree (through shadow roots and slots), the chain of its
 * ancestors, whose styles it inherits, and every style sheet with the path to it (the clone lays
 * ordinary elements out again from those rules, fonts included). The other children of each
 * ancestor stay as empty shells, so sibling selectors and their own sizes still apply; only their
 * contents are skipped. A shell left without its contents can be narrower than it was, so
 * `prepareHtml2CanvasClone` keeps the captured element at its live width. `ignore`, a caller's own
 * predicate, still applies.
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
		const parent = composedParent(node);
		if (parent && ancestors.has(parent)) return false;
		return !(node.closest('head') || holdsStyleSource(node));
	};
}

/**
 * Keep the cloned capture at the live element's layout width. Scoped by `html2canvasCaptureScope`,
 * its ancestors' other children are empty shells, which could narrow a column it takes its width
 * from. Only the width: a fixed height would change how percentage heights inside an auto-height
 * element resolve. The offset width is the layout box, before any transform.
 */
export function pinCapturedSize(original: HTMLElement, clonedEl: HTMLElement): void {
	const width = original.offsetWidth;
	if (width <= 0) return;
	const display = original.ownerDocument.defaultView?.getComputedStyle(original).display;
	if (display === 'inline' || display === 'contents') return;
	Object.assign(clonedEl.style, {
		boxSizing: 'border-box',
		width: `${width}px`,
		minWidth: '0',
		maxWidth: 'none',
	});
}
