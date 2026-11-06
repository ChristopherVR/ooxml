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
 * - a transition layer that the show addresses by id (`preserveElementIds`).
 *   Every binding keeps it on Morph layers, whose per-element keyframe CSS is
 *   keyed on it. Angular also keeps it on its other transition layers, because
 *   they sit inside the show stage and the show applies build states to every
 *   `[data-element-id]` there; vanilla's layers are copies of the presenting
 *   stages and keep it too. Vue and Svelte strip their other transition layers.
 *
 * Every other stage (slides pane and host thumbnails, slide sorter, master and
 * layout previews, presenter previews, reading view, export rasters) is static
 * and carries none in all five bindings: React draws it with its static
 * renderer, vanilla passes `exposeElementIds: false` to `renderSlideStage`, and
 * Vue's and Svelte's `SlideStage` and Angular's `SlideCanvasComponent` decide
 * it with {@link stageKeepsElementIds}.
 *
 * @module render/stage-element-ids
 */

/** What decides whether a stage keeps its element id markers. */
export interface StageElementIdFlags {
	/** The interactive editing canvas. */
	readonly interactive?: boolean;
	/** The live presentation stage. */
	readonly presenting?: boolean;
	/** A transition layer the show addresses by element id (Morph keyframe CSS, build states). */
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
