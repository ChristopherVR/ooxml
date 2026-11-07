/**
 * A slide's structured background gradient
 * (`PptxSlide.backgroundGradientFill`) and the CSS `backgroundGradient` the
 * renderers paint from it.
 *
 * `SlideBuilder.setBackground` stores both. The save path writes the
 * structured gradient only while `backgroundGradient` still is the CSS this
 * module builds from it, so an edit that changed or cleared the CSS (the
 * editor's background controls only know that field) is never overwritten
 * by a stale structured gradient.
 *
 * @module slide-background-gradient
 */

import type { PptxSlide, PptxSlideBackgroundGradient } from '../types';
import { ooxmlGradientAngleToCssDegrees } from './gradient-angle';

/** A number as CSS writes it here: at most three decimals, no trailing zeros. */
const num = (value: number): string => String(Number(value.toFixed(3)));

function cssColor(color: string, opacity: number | undefined): string {
	if (opacity === undefined || opacity >= 1) {
		return color;
	}
	const hex = color.replace('#', '');
	const channel = (index: number) => Number.parseInt(hex.slice(index, index + 2), 16);
	return `rgba(${channel(0)}, ${channel(2)}, ${channel(4)}, ${num(opacity)})`;
}

/** The CSS gradient renderers paint for `gradient`. */
export function slideBackgroundGradientCss(gradient: PptxSlideBackgroundGradient): string {
	const stops = (gradient.fillGradientStops ?? [])
		.map((stop) => `${cssColor(stop.color, stop.opacity)} ${num(stop.position)}%`)
		.join(', ');
	if (gradient.fillGradientType !== 'radial') {
		const angle = ooxmlGradientAngleToCssDegrees(gradient.fillGradientAngle ?? 0);
		return `linear-gradient(${num(angle)}deg, ${stops})`;
	}
	// The centre of the focus rectangle, as the shape gradient codec places it.
	const focus = gradient.fillGradientFillToRect;
	const cx = focus ? ((focus.l + (1 - focus.r)) / 2) * 100 : 50;
	const cy = focus ? ((focus.t + (1 - focus.b)) / 2) * 100 : 50;
	const at = `at ${num(cx)}% ${num(cy)}%`;
	switch (gradient.fillGradientPathType) {
		case 'rect':
			return `radial-gradient(${num(Math.max(cx, 100 - cx))}% ${num(Math.max(cy, 100 - cy))}% ${at}, ${stops})`;
		case 'shape':
			return `radial-gradient(farthest-side ${at}, ${stops})`;
		default:
			return `radial-gradient(circle ${at}, ${stops})`;
	}
}

/**
 * The structured gradient to write for `slide`, or `undefined` when it has
 * none or its CSS gradient no longer is the one built from it.
 */
export function slideBackgroundGradientToWrite(
	slide: Pick<PptxSlide, 'backgroundGradient' | 'backgroundGradientFill'>,
): PptxSlideBackgroundGradient | undefined {
	const gradient = slide.backgroundGradientFill;
	if (!gradient || !slide.backgroundGradient) {
		return undefined;
	}
	return slide.backgroundGradient === slideBackgroundGradientCss(gradient) ? gradient : undefined;
}
