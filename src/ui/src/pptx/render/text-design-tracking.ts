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
	if (!hookInstalled && typeof document !== 'undefined') {
		hookInstalled = true;
		document.fonts?.addEventListener?.('loadingdone', () => paintedCache.clear());
	}
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
}
