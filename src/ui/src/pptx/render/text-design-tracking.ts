import { designTextWidthEm, hasDesignAdvances } from 'ooxml-core/text';

/**
 * Letter-spacing that makes a substituted font take the width of the deck's own font.
 *
 * PowerPoint lays a run out with the deck font's advances. When the browser cannot paint that
 * font (Aptos is a cloud font no browser can read) it paints a stand-in, and a stand-in with
 * other advances (Calibri is 5.6% narrower than Aptos) breaks the lines in other places and
 * sets the text a little smaller (issue #249). The advances of the fonts measured from
 * PowerPoint (`designTextWidthEm`) say what width the run should have, so the difference
 * is spread over its characters. A font the browser does paint is left to the rounding
 * model in `text-metric-tracking.ts`.
 */

/** PowerPoint snaps every glyph advance to a whole pixel at 576 DPI: 6 steps per CSS px. */
const STEPS_PER_PX = 6;

/** Past this the stand-in is a different design, not a width to correct. */
const MAX_TRACKING_EM = 0.15;

const paintedCache = new Map<string, boolean>();
let hookInstalled = false;
let probeContext: CanvasRenderingContext2D | null | undefined;

/**
 * A face that finishes loading can turn a stand-in into the painted deck font, so the answers are
 * dropped. Installed when the module loads, ahead of any binding's own `loadingdone` listener, so
 * a binding that re-renders on that event never reads an answer about the old face.
 */
function installFontLoadHook(): void {
	if (hookInstalled || typeof document === 'undefined') {
		return;
	}
	hookInstalled = true;
	document.fonts?.addEventListener?.('loadingdone', () => paintedCache.clear());
}
installFontLoadHook();

/** The deck font: the first family of the resolved CSS list, unquoted. */
export function deckFamilyOf(fontFamily: string | undefined): string | undefined {
	const first = fontFamily
		?.split(',')[0]
		?.trim()
		.replace(/^["']|["']$/g, '');
	return first || undefined;
}

/**
 * Whether the browser paints `family` itself rather than a stand-in: a string set in
 * `family` followed by one fallback measures differently from the fallback alone only when
 * the family exists, checked against two fallbacks so a family that equals one of them is
 * still seen.
 */
export function isFamilyPainted(ctx: CanvasRenderingContext2D, family: string): boolean {
	const cached = paintedCache.get(family);
	if (cached !== undefined) {
		return cached;
	}
	installFontLoadHook();
	const probe = 'mmmmmmmmmmlliWW1234';
	const width = (font: string): number => {
		ctx.font = font;
		return ctx.measureText(probe).width;
	};
	const painted =
		width(`72px "${family}", monospace`) !== width('72px monospace') ||
		width(`72px "${family}", serif`) !== width('72px serif');
	paintedCache.set(family, painted);
	return painted;
}

/**
 * Whether a run in `fontFamily` (a CSS list, deck font first) takes its width from the deck
 * font's measured advances because the browser paints a stand-in: then its spacing depends on
 * which stand-in is loaded (see `text-tracking-settle.ts`).
 */
export function isDeckFontSubstituted(fontFamily: string | undefined): boolean {
	const deckFamily = deckFamilyOf(fontFamily);
	if (!deckFamily || !hasDesignAdvances(deckFamily)) {
		return false;
	}
	if (probeContext === undefined) {
		probeContext =
			typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
	}
	return probeContext !== null && !isFamilyPainted(probeContext, deckFamily);
}

/**
 * The letter-spacing (CSS px) that gives `text` the width the deck font has in PowerPoint,
 * or `undefined` when the deck font is painted, was not measured, or the text needs
 * a glyph outside the measured range.
 *
 * @param natural The width (px) the browser paints `text` at without any spacing.
 */
export function resolveDesignTrackingPx(
	ctx: CanvasRenderingContext2D,
	text: string,
	deckFamily: string | undefined,
	sizePx: number,
	bold: boolean,
	natural: number,
): number | undefined {
	if (!deckFamily || !hasDesignAdvances(deckFamily) || isFamilyPainted(ctx, deckFamily)) {
		return undefined;
	}
	const chars = [...text];
	let target = 0;
	for (const char of chars) {
		const em = designTextWidthEm(char, deckFamily, bold);
		if (em === undefined) {
			return undefined;
		}
		target += Math.round(em * sizePx * STEPS_PER_PX) / STEPS_PER_PX;
	}
	const limit = MAX_TRACKING_EM * sizePx;
	return Math.min(limit, Math.max(-limit, (target - natural) / chars.length));
}

/** Test hook. */
export function resetDesignTrackingCache(): void {
	paintedCache.clear();
	probeContext = undefined;
}
