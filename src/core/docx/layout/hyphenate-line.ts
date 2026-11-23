import {
	hyphenationPoints,
	SOFT_HYPHEN,
	softHyphenIndices,
	stripSoftHyphens,
} from '../../text/hyphenation';
import type { LayoutHyphenation } from './input';
import type { PlacedToken } from './paragraph-tokens';
import { twips } from '../units';
import { twipsToPx } from './units';

/** Word's default `w:hyphenationZone` when omitted from `settings.xml`: 360 twips (0.25 in = 24 px). */
export const DEFAULT_HYPHENATION_ZONE_PX = twipsToPx(twips(360));

export interface HyphenSplit {
	/** Exact slice of `token.text` consumed on the current line (keeps source offsets accurate). */
	prefixSource: string;
	/** Rendered text for the line-end fragment, ending with a visible `-`. */
	displayText: string;
	/** Measured width of `displayText`. */
	widthPx: number;
	/** Character offset within `token.text` where the next line resumes. */
	resumeOffset: number;
	/** True when chosen by automatic hyphenation (counts toward `consecutiveHyphenLimit`). */
	automatic: boolean;
}

/**
 * Width already used on the current line excluding trailing whitespace that would be trimmed
 * if the line broke before the overflowing word.
 */
export function lineUsedWidthPx(placed: readonly PlacedToken[], lineWidthPx: number): number {
	let trailingSpacePx = 0;
	for (let i = placed.length - 1; i >= 0; i--) {
		const entry = placed[i]!;
		if (entry.token.kind !== 'space') break;
		trailingSpacePx += entry.advancePx ?? entry.widthPx;
	}
	return Math.max(0, lineWidthPx - trailingSpacePx);
}

/**
 * Attempts to split an overflowing word token so its first part plus a visible `-` fits within
 * `availablePx`. Checks explicit soft hyphens (`\u00AD`) first (regardless of `autoHyphenation`),
 * then automatic hyphenation when enabled and permitted by zone and consecutive-hyphen limits.
 */
export function findHyphenSplit(
	word: string,
	availablePx: number,
	gapPx: number,
	hyphenation: LayoutHyphenation | undefined,
	consecutiveHyphens: number,
	measure: (text: string) => number,
): HyphenSplit | undefined {
	if (availablePx <= 0) return undefined;

	// Explicit soft hyphens (`w:softHyphen` / `\u00AD`) are author-placed break points that apply
	// even when automatic hyphenation is off or suppressed.
	if (word.includes(SOFT_HYPHEN)) {
		const softPoints = softHyphenIndices(word);
		for (let i = softPoints.length - 1; i >= 0; i--) {
			const pos = softPoints[i]!;
			const displayText = `${stripSoftHyphens(word.slice(0, pos))}-`;
			const widthPx = measure(displayText);
			if (widthPx <= availablePx) {
				return {
					prefixSource: word.slice(0, pos + 1),
					displayText,
					widthPx,
					resumeOffset: pos + 1,
					automatic: false,
				};
			}
		}
		return undefined;
	}

	if (!hyphenation) return undefined;
	const limit = hyphenation.consecutiveLimit ?? 0;
	if (limit > 0 && consecutiveHyphens >= limit) return undefined;

	const zonePx =
		hyphenation.zoneTwips !== undefined
			? twipsToPx(hyphenation.zoneTwips)
			: DEFAULT_HYPHENATION_ZONE_PX;
	if (gapPx <= zonePx) return undefined;

	const points = hyphenationPoints(
		word,
		hyphenation.doNotHyphenateCaps ? { doNotHyphenateCaps: true } : undefined,
	);
	for (let i = points.length - 1; i >= 0; i--) {
		const pos = points[i]!;
		const prefixSource = word.slice(0, pos);
		const displayText = `${prefixSource}-`;
		const widthPx = measure(displayText);
		if (widthPx <= availablePx) {
			return {
				prefixSource,
				displayText,
				widthPx,
				resumeOffset: pos,
				automatic: true,
			};
		}
	}
	return undefined;
}
