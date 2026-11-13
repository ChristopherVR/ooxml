import { mergeElementRenderErrorScope } from 'ooxml-ui/pptx';
import type { ElementRenderErrorHandler, ElementRenderErrorScope } from 'ooxml-ui/pptx';
import { getContext, setContext } from 'svelte';

/**
 * Carries the host's `onElementRenderError` and the painted slide's id down to every
 * `ElementBoundary`, through the viewer root, the building-blocks slides pane and each
 * `SlideStage` (so thumbnails and the show report their own slide). A getter, because Svelte
 * context is captured once at init and the reporter and slide can change.
 */
export type ElementRenderErrorScopeGetter = () => ElementRenderErrorScope;

/** Exported for out-of-tree mounts and tests: `mount(C, { context: new Map([[key, getter]]) })`. */
export const ElementRenderErrorScopeKey = Symbol('pptx-svelte-element-render-errors');

/** The scope in effect for this component (empty outside a viewer). */
export function getElementRenderErrorScope(): ElementRenderErrorScopeGetter {
	return (
		getContext<ElementRenderErrorScopeGetter | undefined>(ElementRenderErrorScopeKey) ??
		(() => ({}))
	);
}

/** Provide a reporter for this subtree (call during component init). */
export function provideElementRenderErrorReporter(
	reporter: () => ElementRenderErrorHandler | undefined,
): void {
	const inherited = getElementRenderErrorScope();
	setContext<ElementRenderErrorScopeGetter>(ElementRenderErrorScopeKey, () =>
		mergeElementRenderErrorScope(inherited(), { onElementRenderError: reporter() }),
	);
}

/** Name the slide this subtree paints (call during component init). */
export function provideElementRenderSlide(slideId: () => string | undefined): void {
	const inherited = getElementRenderErrorScope();
	setContext<ElementRenderErrorScopeGetter>(ElementRenderErrorScopeKey, () =>
		mergeElementRenderErrorScope(inherited(), { slideId: slideId() }),
	);
}
