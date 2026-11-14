/**
 * The open family that LOOKS most like a font Google Fonts does not serve, for a product that
 * corrects the stand-in's widths itself.
 *
 * `metric-clones.ts` answers a different question: which open family has the same advance widths,
 * so a product that lets the browser measure text keeps its line breaks. Aptos has no such clone;
 * its nearest measured match is Calibri (Carlito), whose letterforms are visibly different (round
 * terminals, a plain `l`, lighter strokes), so Aptos text set in it reads as a different, lighter
 * font even when every line breaks in the right place (issue #249).
 *
 * PowerPoint does not need matching widths from the stand-in: when the deck font is not painted,
 * `text-design-tracking.ts` (`ooxml-ui/pptx`) gives each run the letter-spacing that restores
 * the width PowerPoint measured for the deck font. It can therefore pick the stand-in by shape.
 * Products without that correction keep the metric clone.
 *
 * Measured, not chosen by eye. Aptos and Aptos Display were rendered by PowerPoint (COM
 * `Slide.Export`, 100 pt at 4 px/pt, regular and bold) and compared with about 80 sans-serif
 * Google Fonts families drawn by Chromium at the same size, each word spread to the width
 * PowerPoint gave it, by the overlap (intersection over union) of the inked pixels over sixteen
 * words. Gantari scored 0.762 against Aptos (Calibri 0.650, Carlito 0.665) and 0.724 against
 * Aptos Display (Calibri 0.685, Calibri Light 0.572). The two families that scored marginally
 * higher were rejected: Familjen Grotesk has a single-storey `a`, and Funnel Sans is 17% heavier
 * than Aptos in the regular weight. Gantari is also within 2% of Aptos's x-height and cap height,
 * 3% of its stem and 4% of its ink weight, and only 2.3% wider, so the correction is small.
 *
 * Its bold is heavier than Aptos Bold, though: at slide size (24 pt) Gantari 700 laid down 7.6%
 * more ink than PowerPoint's Aptos Bold, which read as a heavier weight. Gantari is a variable
 * font, so the bold face is requested at 650, which matches Aptos Bold's ink; with only the 400
 * and 650 faces declared, the browser's weight matching paints every bold (700) run with the
 * 650 face, so the runs' own styles stay `font-weight: bold`.
 *
 * The same test, against PowerPoint's renders of each font's words and with every weight of each
 * candidate tried, picked the other look-alikes (intersection over union, regular and bold
 * averaged; the installed original itself scored 0.85 to 0.95, which is the ceiling):
 * Comic Sans MS as Comic Relief (0.821; Comic Neue, the name the CSS chain listed, did not reach
 * the top seven), Consolas as Overpass Mono at 500 and 700 (0.750), Aptos Mono as Sometype Mono at
 * 500 and 700 (0.672), Garamond as EB Garamond at 400 and 600 (0.633) and Impact as Anton (0.749;
 * Anton has one weight). Each of these fonts' own widths are measured through PowerPoint too
 * (`design-advances-powerpoint.generated.ts`), so the stand-in breaks lines where PowerPoint does.
 * Franklin Gothic is left out: its theme pairs Book and Medium, which one look-alike family cannot
 * tell apart by weight while both runs ask for the regular face.
 *
 * @module text/webfonts/look-alikes
 */
import { findGoogleFontsFamily } from './google-fonts-lookup';

interface LookAlike {
	readonly family: string;
	/** The weight whose face regular runs paint with (the browser's weight matching picks it). */
	readonly regularWeight: number;
	/** The weight whose face bold runs paint with: the one that matches the font's bold. */
	readonly boldWeight: number;
}

const LOOK_ALIKES: Readonly<Record<string, LookAlike>> = {
	aptos: { family: 'Gantari', regularWeight: 400, boldWeight: 650 },
	'aptos display': { family: 'Gantari', regularWeight: 400, boldWeight: 650 },
	'aptos mono': { family: 'Sometype Mono', regularWeight: 500, boldWeight: 700 },
	consolas: { family: 'Overpass Mono', regularWeight: 500, boldWeight: 700 },
	'comic sans ms': { family: 'Comic Relief', regularWeight: 400, boldWeight: 700 },
	garamond: { family: 'EB Garamond', regularWeight: 400, boldWeight: 600 },
	impact: { family: 'Anton', regularWeight: 400, boldWeight: 400 },
};

/**
 * The Google Fonts family that looks most like `family`, or `null` when `family` is served under
 * its own name or has no measured look-alike in the catalogue.
 */
export function findLookAlikeGoogleFontsFamily(family: string): string | null {
	if (findGoogleFontsFamily(family) !== null) {
		return null;
	}
	const lookAlike = LOOK_ALIKES[family.trim().toLowerCase()];
	return lookAlike ? findGoogleFontsFamily(lookAlike.family) : null;
}

/** The weights to request for `family`'s look-alike: the faces regular and bold runs paint with. */
export interface LookAlikeWeights {
	readonly regular: number;
	readonly bold: number;
}

/**
 * The weights to request for `family`'s look-alike, or `undefined` when
 * {@link findLookAlikeGoogleFontsFamily} has none for it.
 */
export function findLookAlikeWeights(family: string): LookAlikeWeights | undefined {
	if (findLookAlikeGoogleFontsFamily(family) === null) return undefined;
	const lookAlike = LOOK_ALIKES[family.trim().toLowerCase()];
	return lookAlike ? { regular: lookAlike.regularWeight, bold: lookAlike.boldWeight } : undefined;
}
