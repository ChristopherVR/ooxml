/**
 * Recomputes a run's letter-spacing in place once a web font has loaded.
 *
 * When the deck font is not painted (Aptos is a cloud font), a run gets the letter-spacing that
 * spreads the deck font's measured width over the face the browser does paint
 * (`text-design-tracking.ts`). That spacing depends on the painted face, and the face changes when
 * the webfont fallback finishes loading (the stand-in Calibri gives way to Gantari, issue #249)
 * after the slide was built. Nothing in the five bindings rebuilds a slide when a font loads, so
 * the spacing measured for the old face stayed: wider than the new face needs, and the lines broke
 * early. Such a run carries the {@link TRACKING_PENDING_VAR} custom property with its authored
 * spacing; when fonts finish loading this module measures the face each marked run now paints and
 * rewrites its `letter-spacing` in place. It works on the rendered DOM, like
 * `paragraph-baseline-settle.ts`, so every binding gets it without wiring of its own.
 *
 * @module render/text-tracking-settle
 */
import { isDeckFontSubstituted } from './text-design-tracking';
import { resolveMetricTrackingPx, type RunFontSpec } from './text-metric-tracking';
import { pieceLetterSpacing } from './text-run-spacing';

/** Marks a run whose tracking was measured against a stand-in face; holds the authored px. */
export const TRACKING_PENDING_VAR = '--pptx-tracking-pending';

let armed = false;
let scheduled = false;

/**
 * The marker for a run with this font, or `undefined` when the deck font is painted (or not
 * measured), so its spacing cannot depend on which stand-in is loaded. Arms the settle.
 */
export function trackingPendingStyle(
	font: RunFontSpec,
	authoredPx: number,
): Record<string, string> | undefined {
	if (!isDeckFontSubstituted(font.fontFamily)) {
		return undefined;
	}
	armTrackingSettle();
	return { [TRACKING_PENDING_VAR]: String(authoredPx) };
}

/** The font a marked element paints with, read back from its computed style. */
function paintedFont(element: HTMLElement): RunFontSpec | undefined {
	const style = getComputedStyle(element);
	const fontSizePx = Number.parseFloat(style.fontSize);
	if (!style.fontFamily || !(fontSizePx > 0)) {
		return undefined;
	}
	return {
		fontFamily: style.fontFamily,
		fontSizePx,
		bold: Number.parseInt(style.fontWeight, 10) >= 600,
		italic: style.fontStyle !== 'normal',
	};
}

/** Give every marked run under `root` the spacing for the face it now paints. */
export function settleTracking(root: ParentNode = document): number {
	let settled = 0;
	for (const element of root.querySelectorAll<HTMLElement>(`[style*="${TRACKING_PENDING_VAR}"]`)) {
		const authoredPx = Number.parseFloat(element.style.getPropertyValue(TRACKING_PENDING_VAR));
		const font = paintedFont(element);
		const text = element.textContent ?? '';
		if (!font || !text || !Number.isFinite(authoredPx)) {
			element.style.removeProperty(TRACKING_PENDING_VAR);
			continue;
		}
		const spacing = pieceLetterSpacing(authoredPx, resolveMetricTrackingPx(text, font));
		if (spacing === undefined) {
			element.style.removeProperty('letter-spacing');
		} else {
			element.style.letterSpacing = spacing;
		}
		// Kept while the deck font is still substituted: another face may load later (bold after
		// regular). A run whose deck font is now painted is measured by the rounding model for good.
		if (!isDeckFontSubstituted(font.fontFamily)) {
			element.style.removeProperty(TRACKING_PENDING_VAR);
		}
		settled += 1;
	}
	return settled;
}

function schedule(): void {
	if (scheduled) {
		return;
	}
	scheduled = true;
	const run = () => {
		scheduled = false;
		settleTracking();
	};
	// After the frame: a binding that does re-render on the same event has done so by then.
	if (typeof requestAnimationFrame === 'function') {
		requestAnimationFrame(run);
	} else {
		setTimeout(run, 0);
	}
}

/** Start watching for loaded fonts (once per page). Costs nothing before a run is marked. */
export function armTrackingSettle(): void {
	if (armed || typeof document === 'undefined') {
		return;
	}
	armed = true;
	document.fonts?.addEventListener?.('loadingdone', schedule);
}
