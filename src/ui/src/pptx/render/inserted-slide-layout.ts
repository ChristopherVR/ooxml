/**
 * New Slide from a layout, for every binding: the binding inserts a draft slide (one undo step),
 * then the core re-maps it onto the layout asynchronously (`PptxHandler.applyLayoutToSlide`) and
 * the binding fetches the layout's artwork. By the time both arrive the user (or an automation
 * call) may have edited the new slide; putting the re-mapped draft in its place then would throw
 * that edit away.
 *
 * @module render/inserted-slide-layout
 */

/**
 * What a binding does once a new slide's layout arrives:
 * - `replace`: the draft is still in place, untouched: the re-mapped slide takes its place, and
 *   the layout's artwork goes with it, completing the insertion's step (no step of its own);
 * - `artwork`: the slide was edited meanwhile (the same id, another object): the edit stays, and
 *   only the layout's artwork is taken;
 * - `none`: the slide is gone, or another one is at its index.
 */
export type InsertedLayoutCompletion = 'replace' | 'artwork' | 'none';

/**
 * The completion for a new slide whose draft `placed` was put at an index that now holds
 * `current`. Compares by identity: every edit replaces the slide object.
 */
export function insertedLayoutCompletion(
	current: { readonly id: string } | undefined,
	placed: { readonly id: string },
): InsertedLayoutCompletion {
	if (current === placed) {
		return 'replace';
	}
	return current?.id === placed.id ? 'artwork' : 'none';
}
