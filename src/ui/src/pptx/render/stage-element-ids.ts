/**
 * Which slide stages keep the `data-element-id` markers, and the DOM pass that
 * removes them from the others.
 *
 * Every element renderer stamps `data-element-id`, but the marker must be
 * unique to the surface a user acts on: hit-testing, animation preview, paste
 * anchoring and `[data-element-id]` queries take the first match, and the
 * thumbnail rail sits ahead of the canvas in the DOM. So only three kinds of
 * stage keep it:
 *
 * - the interactive editing canvas;
 * - the live presentation stage (playback, triggers and builds select by it);
 * - a Morph transition layer, whose per-element keyframe CSS is keyed on it
 *   (the host hides its own stage while the overlay is mounted, so there is
 *   still one marked copy of each element).
 *
 * Every other stage (slides pane and host thumbnails, slide sorter, master and
 * layout previews, presenter previews, reading view, non-Morph transition
 * layers, export rasters) is static and carries none, as React's static
 * renderer and Angular's `exposeElementIds` already do.
 *
 * @module render/stage-element-ids
 */

/** What decides whether a stage keeps its element id markers. */
export interface StageElementIdFlags {
	/** The interactive editing canvas. */
	readonly interactive?: boolean;
	/** The live presentation stage. */
	readonly presenting?: boolean;
	/** A Morph transition layer, whose keyframe CSS selects elements by id. */
	readonly preserveElementIds?: boolean;
}

/** Whether a stage with these flags keeps its `data-element-id` markers. */
export function stageKeepsElementIds(flags: StageElementIdFlags): boolean {
	return Boolean(flags.interactive || flags.presenting || flags.preserveElementIds);
}

/**
 * Remove every `data-element-id` marker under `root` (and from `root` itself)
 * and return how many were removed. Content and every other attribute stay.
 * Run it after the stage is assembled: renderers that index their own subtree
 * by id (group fills) need the markers while they build.
 */
export function stripElementIdMarkers(root: ParentNode): number {
	const nodes = [...root.querySelectorAll('[data-element-id]')];
	if (
		typeof Element !== 'undefined' &&
		root instanceof Element &&
		root.hasAttribute('data-element-id')
	) {
		nodes.push(root);
	}
	for (const node of nodes) {
		node.removeAttribute('data-element-id');
	}
	return nodes.length;
}
