import { definedProps } from '../defined-props';
import { stripSoftHyphens } from '../../text/hyphenation';
import type { LayoutFragment } from './result';
import type { LayoutParagraph } from './input';
import type { PlacedToken } from './paragraph-tokens';

export function buildFragments(
	tokensOnLine: PlacedToken[],
	paragraph: LayoutParagraph,
): LayoutFragment[] {
	const fragments: LayoutFragment[] = [];
	let x = 0;
	for (const {
		token,
		widthPx,
		displayText,
		leader,
		xOffsetPx,
		advancePx,
		kernAfterPx,
	} of tokensOnLine) {
		const run = paragraph.runs[token.runIndex];
		const text =
			displayText ??
			(token.kind === 'word' || token.kind === 'space' ? stripSoftHyphens(token.text) : '');
		fragments.push({
			text,
			xPx: x + (xOffsetPx ?? 0),
			widthPx,
			runIndex: token.runIndex,
			...definedProps({
				bold: run?.bold,
				italic: run?.italic,
				fontFamily: run?.fontFamily,
				fontFallback: run?.fontFallback,
				fontSizePt: run?.fontSizePt,
				contentControlTag: run?.contentControlTag,
				blockContentControlTag: run?.blockContentControlTag,
				textScalePercent: run?.textScalePercent,
				ligatures: run?.ligatures,
				smallCaps: run?.smallCaps,
				characterSpacingPx: run?.characterSpacingPx,
				kerningThresholdPt: run?.kerningThresholdPt,
			}),
			...(kernAfterPx ? { kernAfterPx } : {}),
			...(leader ? { leader } : {}),
			...(run?.script ? { script: run.script } : {}),
			...(run?.color ? { color: run.color } : {}),
			...(run?.background ? { background: run.background } : {}),
			...(run?.mergeField ? { mergeField: true } : {}),
			...(run?.underline ? { underline: true } : {}),
			...(run?.strike ? { strike: true } : {}),
			...(token.kind === 'object' && run?.object ? { object: run.object } : {}),
		});
		x += advancePx ?? widthPx;
	}
	return fragments;
}

function justify(fragments: LayoutFragment[], width: number): LayoutFragment[] {
	const spaceIndices = fragments.map((f, i) => (f.text === ' ' ? i : -1)).filter((i) => i >= 0);
	if (!spaceIndices.length || !fragments.length) return fragments;
	const natural = fragments.at(-1)!.xPx + fragments.at(-1)!.widthPx;
	// A line that was fitted by squeezing its spaces has a negative remainder.
	const extra = (width - natural) / spaceIndices.length;
	if (extra === 0) return fragments;
	let shift = 0;
	return fragments.map((fragment, index) => {
		const placedFragment = { ...fragment, xPx: fragment.xPx + shift };
		if (spaceIndices.includes(index)) {
			placedFragment.widthPx += extra;
			shift += extra;
		}
		return placedFragment;
	});
}

export function alignFragments(
	fragments: LayoutFragment[],
	width: number,
	skipExpansion: boolean,
	effectiveAlign: 'left' | 'right' | 'center' | 'justify',
	squeezed = false,
): LayoutFragment[] {
	if (!fragments.length) return fragments;
	const natural = fragments.at(-1)!.xPx + fragments.at(-1)!.widthPx;
	if (effectiveAlign === 'justify' && (!skipExpansion || (squeezed && natural > width)))
		return justify(fragments, width);
	const offset =
		effectiveAlign === 'center'
			? Math.max(0, width - natural) / 2
			: effectiveAlign === 'right'
				? Math.max(0, width - natural)
				: 0;
	return offset === 0 ? fragments : fragments.map((f) => ({ ...f, xPx: f.xPx + offset }));
}
