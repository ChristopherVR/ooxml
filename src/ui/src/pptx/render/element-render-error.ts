/**
 * Per-element render error containment, shared by the five PowerPoint bindings.
 *
 * A slide element whose render or lifecycle throws is omitted on its own: its neighbours, the
 * slide and the viewer chrome stay. Each binding keeps only its native boundary (React error
 * boundary, Vue `onErrorCaptured`, Svelte `<svelte:boundary>`, Angular's element host, vanilla's
 * try/catch in the stage) and takes every decision from here: what the host is told
 * ({@link elementRenderErrorContext}), how it is told without a throwing reporter taking the
 * viewer down ({@link reportElementRenderError}), and when a failed element is tried again
 * ({@link shouldRetryElementRender}: a replacement element object or a changed id).
 *
 * @module render/element-render-error
 */
import type { PptxElement } from 'ooxml-core/pptx';

/** Identifies the element omitted after a render or lifecycle error. */
export interface ElementRenderErrorContext {
	elementId: string;
	elementType: PptxElement['type'];
	/** The slide that owns the element, when the surface knows it. */
	slideId?: string;
}

/** The host reporter every binding accepts as `onElementRenderError`. */
export type ElementRenderErrorHandler = (error: Error, context: ElementRenderErrorContext) => void;

/** The reporting scope a binding carries down to its element boundaries (context, injection). */
export interface ElementRenderErrorScope {
	slideId?: string | undefined;
	onElementRenderError?: ElementRenderErrorHandler | undefined;
}

/** The identity a failed element is remembered by. */
export interface ElementRenderRetryKey {
	readonly element: PptxElement;
	readonly id: string;
}

/** Remember `element` as failed. */
export function elementRenderRetryKey(element: PptxElement): ElementRenderRetryKey {
	return { element, id: element.id };
}

/**
 * Whether a boundary should render `element`: nothing failed yet, or the failed element was
 * replaced by another object or its id changed. Rendering the same failed object again (an
 * in-place mutation included) does not retry.
 */
export function shouldRetryElementRender(
	failed: ElementRenderRetryKey | null | undefined,
	element: PptxElement,
): boolean {
	return !failed || failed.element !== element || failed.id !== element.id;
}

/** The payload reported for `element`, with the owning slide when one is known. */
export function elementRenderErrorContext(
	element: PptxElement,
	slideId?: string | undefined,
): ElementRenderErrorContext {
	return {
		elementId: element.id,
		elementType: element.type,
		...(slideId === undefined ? {} : { slideId }),
	};
}

/** Errors are passed through unchanged; any other thrown value is wrapped. */
export function toElementRenderError(thrown: unknown): Error {
	return thrown instanceof Error ? thrown : new Error(String(thrown));
}

/**
 * Tell the host about an omitted element. Without a reporter, or when the reporter itself
 * throws, the error goes to `console.error(error, context)`; this never throws.
 */
export function reportElementRenderError(
	handler: ElementRenderErrorHandler | undefined,
	thrown: unknown,
	context: ElementRenderErrorContext,
): void {
	const error = toElementRenderError(thrown);
	if (handler) {
		try {
			handler(error, context);
			return;
		} catch {
			// A host reporter must not turn an isolated render error into a viewer crash.
		}
	}
	try {
		console.error(error, context);
	} catch {
		// Nothing left to report to.
	}
}

/** A nested scope: what the inner surface sets wins, the rest is inherited. */
export function mergeElementRenderErrorScope(
	inherited: ElementRenderErrorScope | undefined,
	own: ElementRenderErrorScope,
): ElementRenderErrorScope {
	return {
		slideId: own.slideId ?? inherited?.slideId,
		onElementRenderError: own.onElementRenderError ?? inherited?.onElementRenderError,
	};
}

/**
 * The elements a surface has seen fail, for bindings that rebuild their DOM instead of keeping a
 * boundary per element (vanilla). A failed element is skipped, and not reported again, until
 * {@link shouldRetryElementRender} says it changed.
 */
export interface ElementRenderFailures {
	/** True when `element` failed before and has not changed since. */
	has(element: PptxElement): boolean;
	/** Remember `element` as failed. */
	add(element: PptxElement): void;
}

export function createElementRenderFailures(): ElementRenderFailures {
	const failed = new WeakMap<PptxElement, string>();
	return {
		has: (element) => failed.get(element) === element.id,
		add: (element) => {
			failed.set(element, element.id);
		},
	};
}

export interface GuardElementRenderOptions extends ElementRenderErrorScope {
	/** Failures remembered across rebuilds; without it every call tries again. */
	failures?: ElementRenderFailures | undefined;
	/** Undo anything the failed attempt left behind (detached nodes, subscriptions). */
	onFailure?: (() => void) | undefined;
}

/**
 * Build one element imperatively, returning `null` (and reporting) instead of throwing when
 * the build fails. The try/catch boundary of a binding without component boundaries.
 */
export function guardElementRender<T>(
	element: PptxElement,
	build: () => T,
	options: GuardElementRenderOptions = {},
): T | null {
	if (options.failures?.has(element)) {
		return null;
	}
	try {
		probeElementRenderFault(element);
		return build();
	} catch (thrown) {
		options.failures?.add(element);
		try {
			options.onFailure?.();
		} catch {
			// Cleanup is best effort; the element is omitted either way.
		}
		reportElementRenderError(
			options.onElementRenderError,
			thrown,
			elementRenderErrorContext(element, options.slideId),
		);
		return null;
	}
}

/** The global a test page sets to the id of the element whose render should throw. */
export const ELEMENT_RENDER_FAULT_KEY = 'ooxml-ui.pptx.element-render-fault';

/**
 * @internal Test hook, not an API: throws when the page set the global
 * `Symbol.for(ELEMENT_RENDER_FAULT_KEY)` to this element's id. The demos set it from the
 * `?elementRenderFault=<id>` query parameter so the browser suite can prove each binding omits
 * only the failing element. Every binding's element renderer calls it once per render.
 */
export function probeElementRenderFault(element: PptxElement): void {
	const fault = (globalThis as Record<symbol, unknown>)[Symbol.for(ELEMENT_RENDER_FAULT_KEY)];
	if (fault !== undefined && fault === element.id) {
		throw new Error(`Injected element render fault (${element.id})`);
	}
}
