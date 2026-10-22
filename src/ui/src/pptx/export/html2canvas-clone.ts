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
	snapshotThreeViewsIntoClone(original, clonedEl);
	prepareExportClone(clonedEl);
	await normalizeColorsForCapture(doc, clonedEl);
	preprocessCssForCapture(clonedEl);
}

/** One step up the composed tree: a node's parent, or the host of the shadow root it sits in. */
function composedParent(node: Node): Node | null {
	const parent = node.parentNode;
	return parent && 'host' in parent ? (parent as ShadowRoot).host : parent;
}

/**
 * The `ignoreElements` every binding hands html2canvas-pro.
 *
 * html2canvas-pro clones the whole document, every element and open shadow root with its computed
 * style copied, before it paints the one element it captures. Under a viewer with its chrome
 * mounted that clone is most of a capture (seconds, all on the main thread), and nothing outside
 * the captured element is painted. This keeps the captured element's subtree (through shadow
 * roots), the chain of its ancestors, whose styles it inherits, and the head and style sheets the
 * clone's fonts and rules come from; it skips every other branch. Each cloned element keeps its own
 * computed size, so leaving out its neighbours does not move or resize what is painted. `ignore`, a
 * caller's own predicate, still applies.
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
		return !(node.closest('head') || node.localName === 'style' || node.localName === 'link');
	};
}
