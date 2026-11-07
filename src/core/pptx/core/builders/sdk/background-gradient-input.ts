/**
 * Map `SlideBuilder.setBackground({ type: 'gradient' })` onto the slide's
 * structured background gradient (written as `p:bg/p:bgPr/a:gradFill` by the
 * shape gradient writer) and the CSS gradient renderers paint.
 *
 * @module sdk/background-gradient-input
 */

import type { PptxSlide, PptxSlideBackgroundGradient } from '../../types/presentation';
import { cssDegreesToOoxmlGradientAngle } from '../../utils/gradient-angle';
import { slideBackgroundGradientCss } from '../../utils/slide-background-gradient';
import type { BackgroundInput } from './types';

type GradientBackgroundInput = Extract<BackgroundInput, { type: 'gradient' }>;

const HEX = /^#?[0-9a-f]{6}$/iu;
const PATHS: ReadonlySet<string> = new Set(['circle', 'rect', 'shape']);
const WHERE = 'BackgroundInput';

function finite(name: string, value: unknown, min?: number, max?: number): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) {
		throw new RangeError(`${WHERE}.${name} must be a finite number, got ${String(value)}`);
	}
	if ((min !== undefined && value < min) || (max !== undefined && value > max)) {
		throw new RangeError(`${WHERE}.${name} must be between ${min} and ${max}, got ${value}`);
	}
	return value;
}

function stopsOf(
	input: GradientBackgroundInput,
): NonNullable<PptxSlideBackgroundGradient['fillGradientStops']> {
	if (!Array.isArray(input.stops) || input.stops.length < 2) {
		throw new RangeError(`${WHERE}.stops needs at least two stops (a:gsLst)`);
	}
	return input.stops.map((stop, index) => {
		if (typeof stop.color !== 'string' || !HEX.test(stop.color)) {
			throw new RangeError(
				`${WHERE}.stops[${index}].color must be #RRGGBB, got ${String(stop.color)}`,
			);
		}
		const position = finite(`stops[${index}].position`, stop.position, 0, 1);
		const opacity =
			stop.opacity === undefined
				? undefined
				: finite(`stops[${index}].opacity`, stop.opacity, 0, 1);
		return {
			color: `#${stop.color.replace('#', '').toUpperCase()}`,
			// The shape gradient model keeps positions in percent.
			position: position * 100,
			...(opacity !== undefined ? { opacity } : {}),
		};
	});
}

/**
 * The slide fields for a gradient background: the structured gradient and
 * its CSS. A linear gradient defaults to `180` (CSS, top to bottom); a path
 * gradient grows from the centre of `focus` (default: the slide centre).
 *
 * @throws RangeError for fewer than two stops, a stop colour that is not
 *   `#RRGGBB`, a position or opacity outside 0-1, an unknown `path`, an
 *   `angle` with a `path`, or a `focus` without one.
 */
export function mapBackgroundGradientInput(
	input: GradientBackgroundInput,
): Pick<PptxSlide, 'backgroundGradient' | 'backgroundGradientFill'> {
	const fillGradientStops = stopsOf(input);
	let gradient: PptxSlideBackgroundGradient;
	if (input.path !== undefined) {
		if (!PATHS.has(input.path)) {
			throw new RangeError(
				`${WHERE}.path must be circle, rect or shape, got ${String(input.path)}`,
			);
		}
		if (input.angle !== undefined) {
			throw new RangeError(`${WHERE}.angle applies to a linear gradient, not a path one`);
		}
		// Without a focus the gradient starts at the centre, written out as
		// PowerPoint does (an absent `a:fillToRect` would mean the whole slide).
		const focus = input.focus ?? { l: 0.5, t: 0.5, r: 0.5, b: 0.5 };
		gradient = {
			fillGradientType: 'radial',
			fillGradientPathType: input.path,
			fillGradientStops,
			fillGradientFillToRect: {
				l: finite('focus.l', focus.l),
				t: finite('focus.t', focus.t),
				r: finite('focus.r', focus.r),
				b: finite('focus.b', focus.b),
			},
		};
	} else {
		if (input.focus !== undefined) {
			throw new RangeError(`${WHERE}.focus needs a path gradient`);
		}
		const cssAngle = finite('angle', input.angle ?? 180);
		gradient = {
			fillGradientType: 'linear',
			fillGradientAngle: cssDegreesToOoxmlGradientAngle(cssAngle),
			fillGradientStops,
		};
	}
	return {
		backgroundGradient: slideBackgroundGradientCss(gradient),
		backgroundGradientFill: gradient,
	};
}
