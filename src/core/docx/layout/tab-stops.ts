import type { LayoutRun, LayoutTabStop } from './input';
import { DEFAULT_TAB_STOP_PX } from './units';

export interface TabPlacement {
	widthPx: number;
	/** The text after the tab is placed relative to the stop (right, centre or decimal). */
	aligned?: true;
	leader?: LayoutTabStop['leader'];
}

/**
 * Where a tab at `positionPx` (from the paragraph's text margin) advances to, as Word does:
 * the next custom stop after it, else the hanging indent (an implicit stop for list labels),
 * else the next default stop. `followingPx` is the width of the text after the tab up to the
 * next tab or line end, used by right, center and decimal stops; `beforeDecimalPx` is the part
 * of it before the first decimal separator. A right stop past `rightEdgePx` (the line's right
 * limit) acts as if it were at that edge: Word ends the text there rather than wrapping it.
 */
export function placeTab(
	positionPx: number,
	stops: readonly LayoutTabStop[],
	hangingStopPx: number | undefined,
	followingPx: number,
	beforeDecimalPx: number,
	defaultStopPx = DEFAULT_TAB_STOP_PX,
	rightEdgePx = Infinity,
): TabPlacement {
	const epsilon = 0.5;
	const custom = stops
		.filter((stop) => stop.align !== 'clear' && stop.align !== 'bar')
		.sort((a, b) => a.posPx - b.posPx)
		.find((stop) => stop.posPx > positionPx + epsilon);
	if (custom) {
		const rightAligned = custom.align === 'right' || custom.align === 'end';
		const stopPx =
			rightAligned && custom.posPx > rightEdgePx && positionPx < rightEdgePx
				? rightEdgePx
				: custom.posPx;
		const distance = stopPx - positionPx;
		const shift = rightAligned
			? followingPx
			: custom.align === 'center'
				? followingPx / 2
				: custom.align === 'decimal'
					? beforeDecimalPx
					: 0;
		return {
			widthPx: Math.max(0, distance - shift),
			...(rightAligned || custom.align === 'center' || custom.align === 'decimal'
				? { aligned: true as const }
				: {}),
			...(custom.leader && custom.leader !== 'none' ? { leader: custom.leader } : {}),
		};
	}
	if (hangingStopPx !== undefined && hangingStopPx > positionPx + epsilon)
		return { widthPx: hangingStopPx - positionPx };
	const next = (Math.floor(positionPx / defaultStopPx) + 1) * defaultStopPx;
	return { widthPx: next - positionPx };
}

/**
 * Where a positional tab (`w:ptab`) at `positionPx` advances to, whatever the tab stops: its
 * alignment point between `startPx` and `endPx` (the text margins or the paragraph's indents),
 * with the text after it (`followingPx` wide) starting, centred or ending there. Undefined when
 * the line is already past that point, so the tab goes to it on the next line.
 */
export function placePositionalTab(
	positionPx: number,
	tab: NonNullable<LayoutRun['positionalTab']>,
	startPx: number,
	endPx: number,
	followingPx: number,
): TabPlacement | undefined {
	const target =
		tab.alignment === 'left' ? startPx : tab.alignment === 'center' ? (startPx + endPx) / 2 : endPx;
	if (positionPx > target + 0.5) return undefined;
	const shift =
		tab.alignment === 'right' ? followingPx : tab.alignment === 'center' ? followingPx / 2 : 0;
	return {
		widthPx: Math.max(0, target - shift - positionPx),
		...(tab.alignment === 'left' ? {} : { aligned: true as const }),
		...(tab.leader && tab.leader !== 'none' ? { leader: tab.leader } : {}),
	};
}
