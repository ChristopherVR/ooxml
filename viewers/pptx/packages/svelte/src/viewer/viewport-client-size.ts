import { observeViewportClientSize } from 'ooxml-ui/pptx';
import type { ViewportClientSize } from 'ooxml-ui/pptx';

/**
 * Svelte action: report the viewport's fractional client size now and on every resize.
 *
 * Replaces `bind:clientWidth` / `bind:clientHeight` on the fit viewport. Those are whole
 * pixels, rounded up by up to half a pixel, and a slide fitted to them overflows the viewport
 * by that much. The measurement itself is the shared `observeViewportClientSize`, which every
 * binding uses.
 */
export function viewportClientSize(
	node: HTMLElement,
	initial: (size: ViewportClientSize) => void,
): { update(next: (size: ViewportClientSize) => void): void; destroy(): void } {
	let onSize = initial;
	const stop = observeViewportClientSize(node, (size) => onSize(size));
	return {
		update(next) {
			onSize = next;
		},
		destroy: stop,
	};
}
