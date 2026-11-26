import type { PptxAnimationDirection } from '../types';
import { getCapturedTree } from './animation-behavior-captured';
import { DIRECTION_TO_SUBTYPE } from './animation-write-mappings-motion';
import type { OoxmlPresetMapping } from './animation-write-mappings-types';

type Variants = Partial<Record<PptxAnimationDirection, number>>;

const SHAPE: Variants = { in: 16, out: 32 };
const ZOOM_IN: Variants = {
	...SHAPE,
	objectCenter: 16,
	inSlightly: 272,
	outSlightly: 288,
	inFromScreenCenter: 528,
	outFromScreenBottom: 36,
};
const ZOOM_OUT: Variants = {
	...SHAPE,
	objectCenter: 32,
	inSlightly: 272,
	outSlightly: 288,
	outFromScreenCenter: 544,
	inToScreenBottom: 20,
};
const VARIANTS: Partial<Record<number, Variants>> = {
	3: { horizontal: 10, vertical: 5 },
	4: SHAPE,
	6: SHAPE,
	8: SHAPE,
	13: SHAPE,
	16: { horizontalIn: 26, horizontalOut: 42, verticalIn: 21, verticalOut: 37 },
};

/** Translate an Effect Options choice into the subtype PowerPoint writes. */
export function resolveDirectionSubtype(
	mapping: OoxmlPresetMapping,
	direction: PptxAnimationDirection | undefined,
): number {
	const { presetClass, presetId, defaultSubtype } = mapping;
	if (presetClass === 'entr' || presetClass === 'exit') {
		if (presetId === 53) {
			// PowerPoint COM writes the same center subtypes as Zoom.
			if (direction === 'objectCenter') return presetClass === 'entr' ? 16 : 32;
			if (direction === 'slideCenter') return presetClass === 'entr' ? 528 : 544;
			return defaultSubtype;
		}
		const variants =
			presetId === 23 ? (presetClass === 'entr' ? ZOOM_IN : ZOOM_OUT) : VARIANTS[presetId];
		const requested = direction
			? variants
				? variants[direction]
				: DIRECTION_TO_SUBTYPE[direction]
			: undefined;
		return (
			getCapturedTree(presetClass, presetId, requested)?.subtype ?? requested ?? defaultSubtype
		);
	}
	return defaultSubtype;
}
