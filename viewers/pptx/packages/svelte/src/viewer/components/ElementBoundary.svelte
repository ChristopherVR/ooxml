<script lang="ts">
	/**
	 * ElementBoundary: omits one slide element whose render or effect throws, so its neighbours,
	 * the slide and the viewer stay (a Svelte 5 `<svelte:boundary>` around the element). The host
	 * hears about it through the context's `onElementRenderError` (default `console.error`). A
	 * replacement element object or a changed id renders it again; re-rendering the same failed
	 * object does not. Event handler errors are not caught, as in React's boundaries.
	 */
	import type { Snippet } from 'svelte';
	import type { PptxElement } from 'pptx-viewer-core';
	import {
		elementRenderErrorContext,
		elementRenderRetryKey,
		reportElementRenderError,
		shouldRetryElementRender,
	} from 'ooxml-ui/pptx';
	import type { ElementRenderRetryKey } from 'ooxml-ui/pptx';

	import { getElementRenderErrorScope } from '../state/element-render-errors';

	const { element, children }: { element: PptxElement; children: Snippet } = $props();

	const scope = getElementRenderErrorScope();
	// Plain variables on purpose: written by the boundary's error callback, read by the effect
	// below, which must track the element alone.
	let failed: ElementRenderRetryKey | null = null;
	let reset: (() => void) | null = null;

	function onerror(error: unknown, retry: () => void): void {
		failed = elementRenderRetryKey(element);
		reset = retry;
		const { slideId, onElementRenderError } = scope();
		reportElementRenderError(onElementRenderError, error, elementRenderErrorContext(element, slideId));
	}

	// A replacement object (or a changed id) tries the element again.
	$effect.pre(() => {
		const current = element;
		void current.id;
		if (reset && shouldRetryElementRender(failed, current)) {
			const retry = reset;
			failed = null;
			reset = null;
			retry();
		}
	});
</script>

<svelte:boundary {onerror}>
	{@render children()}
</svelte:boundary>
