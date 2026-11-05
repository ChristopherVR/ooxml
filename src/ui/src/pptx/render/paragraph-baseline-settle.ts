/**
 * Places a paragraph on PowerPoint's baseline again once its web font has
 * loaded.
 *
 * A baseline move is computed from the font's own ascent and descent
 * (`paragraph-baseline`). A web font (the Google Fonts fallback for a family
 * the machine lacks) has not loaded when a slide is first built, and nothing
 * in the five bindings rebuilds the slide when it does, so the move was
 * computed from the stand-in face and stayed wrong: Oswald and Playfair
 * Display titles sat 4 to 10 px low at 72pt. Such a paragraph carries the
 * `--pptx-baseline-pending` custom property with what is needed to place it
 * again. This module finds the marked paragraphs in the document when fonts
 * finish loading (and when the DOM changes, since a slide switch mounts
 * fresh elements), measures the font that is now painted and rewrites the
 * `top` in place, then drops the marker. It works on the rendered DOM so
 * every binding gets it without wiring of its own.
 *
 * @module render/paragraph-baseline-settle
 */
import {
	BASELINE_PENDING_VAR,
	decodePendingBaseline,
	resolveParagraphBaseline,
} from './paragraph-baseline';

let armed = false;
let scheduled = false;

/** Place every marked paragraph under `root` on the baseline of the font it now paints with. */
export function settleBaselines(root: ParentNode = document): number {
	let remaining = 0;
	for (const element of root.querySelectorAll<HTMLElement>(`[style*="${BASELINE_PENDING_VAR}"]`)) {
		const input = decodePendingBaseline(element.style.getPropertyValue(BASELINE_PENDING_VAR));
		const resolved = input ? resolveParagraphBaseline(input) : undefined;
		if (!resolved) {
			element.style.removeProperty(BASELINE_PENDING_VAR);
			continue;
		}
		if (resolved.shift === undefined) {
			element.style.removeProperty('top');
		} else {
			element.style.position = 'relative';
			element.style.top = `${resolved.shift}px`;
		}
		if (resolved.provisional) {
			remaining += 1;
		} else {
			element.style.removeProperty(BASELINE_PENDING_VAR);
		}
	}
	return remaining;
}

function schedule(): void {
	if (scheduled) {
		return;
	}
	scheduled = true;
	const run = () => {
		scheduled = false;
		settleBaselines();
	};
	if (typeof requestAnimationFrame === 'function') {
		requestAnimationFrame(run);
	} else {
		setTimeout(run, 0);
	}
}

/**
 * Start watching for loaded fonts and new paragraphs (once per page). Call it
 * when a paragraph was marked pending; it costs nothing before that.
 */
export function armBaselineSettle(): void {
	if (armed || typeof document === 'undefined') {
		return;
	}
	armed = true;
	document.fonts?.addEventListener?.('loadingdone', schedule);
	if (typeof MutationObserver === 'function') {
		new MutationObserver(schedule).observe(document.documentElement, {
			childList: true,
			subtree: true,
		});
	}
	schedule();
}
