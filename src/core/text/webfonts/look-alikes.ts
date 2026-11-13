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
 * @module text/webfonts/look-alikes
 */
import { findGoogleFontsFamily } from './google-fonts-lookup';

const LOOK_ALIKES: Readonly<Record<string, string>> = {
	aptos: 'Gantari',
	'aptos display': 'Gantari',
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
	return lookAlike ? findGoogleFontsFamily(lookAlike) : null;
}
