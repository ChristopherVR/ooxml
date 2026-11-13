/**
 * Demo-only test hook for `e2e/pptx/element-render-errors.spec.ts`: `?elementRenderFault=<id>`
 * makes the renderer of the slide element with that id throw, in every binding, through the
 * shared `probeElementRenderFault` (`ooxml-ui/pptx`), which reads the global set here. Not an API:
 * without the query parameter nothing is set and the probe does nothing.
 *
 * Imported for its side effect by each demo's entry, before the viewer mounts.
 */
const raw =
	typeof window === 'undefined'
		? null
		: new URLSearchParams(window.location.search).get('elementRenderFault');
if (raw) {
	(globalThis as Record<symbol, unknown>)[Symbol.for('ooxml-ui.pptx.element-render-fault')] = raw;
}

export {};
