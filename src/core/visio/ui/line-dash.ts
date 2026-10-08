import type { VisioStyle } from '../model';

/** Native dash lengths in physical page inches, including fixed square-cap dots. */
export function visioLineDashLengths(
	style: Pick<VisioStyle, 'lineDash' | 'lineWidth' | 'lineDashDotLength'>,
): number[] | undefined {
	return style.lineDash?.map((length, index) =>
		length === 0 && index % 2 === 0 ? (style.lineDashDotLength ?? 0) : length * style.lineWidth,
	);
}
