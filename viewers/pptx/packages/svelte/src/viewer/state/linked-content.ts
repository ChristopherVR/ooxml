import { getContext, setContext } from 'svelte';

/**
 * Whether this viewer draws the deck's linked web content (the shared
 * `resolveExternalContentView(...).allowed`). Every surface that draws the model
 * reads it through `linkedContentView(value, getLinkedContent())`, which withholds
 * the linked addresses until the viewer allows them; Enable Content changes the
 * value and the surfaces redraw. Nothing reloads and the model never changes.
 *
 * `createViewerState` provides it for the viewer and the building blocks;
 * `SlidesPane` and `SlideThumbnail` re-provide it from their `linkedContent` prop.
 * Outside both (a bare stage, a test) linked content is withheld.
 *
 * The provided value is a getter: Svelte context is captured at component init,
 * and reading the getter inside a `$derived` keeps the surface live.
 */
export type LinkedContentGetter = () => boolean;

/** Exported for out-of-tree mounts (the export capture stage) and tests. */
export const LinkedContentKey = Symbol('pptx-svelte-linked-content');

const WITHHELD: LinkedContentGetter = () => false;

/** Provide the getter to the component subtree. */
export function provideLinkedContent(getAllowed: LinkedContentGetter): void {
	setContext(LinkedContentKey, getAllowed);
}

/** The getter for this subtree; withholds when nothing provides one. Call at component init. */
export function getLinkedContentGetter(): LinkedContentGetter {
	return getContext<LinkedContentGetter | undefined>(LinkedContentKey) ?? WITHHELD;
}

/**
 * Re-provide the getter from a component's optional `linkedContent` prop, falling back to the
 * enclosing context for an unset prop. Call at component init.
 */
export function provideLinkedContentProp(getProp: () => boolean | undefined): void {
	const inherited = getLinkedContentGetter();
	provideLinkedContent(() => getProp() ?? inherited());
}
