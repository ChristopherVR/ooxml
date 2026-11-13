import { mergeElementRenderErrorScope } from 'ooxml-ui/pptx';
import type { ElementRenderErrorHandler, ElementRenderErrorScope } from 'ooxml-ui/pptx';
import type { InjectionKey } from 'vue';
import { inject, provide } from 'vue';

/**
 * Carries the host's `onElementRenderError` and the painted slide's id down to every
 * `ElementBoundary`, through the viewer root, the building-blocks canvas and slides pane, and each
 * `SlideStage` (so thumbnails and the show report their own slide). A getter, so a boundary reads
 * the latest reporter when an element fails.
 */
const ElementRenderErrorScopeKey: InjectionKey<() => ElementRenderErrorScope> = Symbol(
	'pptx-vue-element-render-errors',
);

/** The scope in effect for this component (empty outside a viewer). */
export function injectElementRenderErrorScope(): () => ElementRenderErrorScope {
	return inject(ElementRenderErrorScopeKey, () => ({}));
}

/** Provide a reporter for this subtree (must be called from `setup`). */
export function provideElementRenderErrorReporter(
	reporter: () => ElementRenderErrorHandler | undefined,
): void {
	const inherited = injectElementRenderErrorScope();
	provide(ElementRenderErrorScopeKey, () =>
		mergeElementRenderErrorScope(inherited(), { onElementRenderError: reporter() }),
	);
}

/** Name the slide this subtree paints (must be called from `setup`). */
export function provideElementRenderSlide(slideId: () => string | undefined): void {
	const inherited = injectElementRenderErrorScope();
	provide(ElementRenderErrorScopeKey, () =>
		mergeElementRenderErrorScope(inherited(), { slideId: slideId() }),
	);
}

/**
 * Whether an error Vue captured comes from an event handler. Those are not render or lifecycle
 * failures (React's boundaries do not catch them either), so they keep propagating. Vue names the
 * source in development and passes an error-reference URL with its code in production.
 */
export function isEventHandlerErrorInfo(info: string): boolean {
	return /event handler$/u.test(info) || /#runtime-(?:5|6)$/u.test(info);
}
