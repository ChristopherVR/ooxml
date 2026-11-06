/**
 * Gradients CSS cannot express, painted as a small self-contained SVG image a
 * `background-image: url(...)` can take (the same route `path-gradient-rect.ts`
 * uses for `path="rect"`).
 *
 * Both were measured against PowerPoint 16 (COM `Slide.Export`):
 *
 *  - `a:lin scaled="1"` on a non-square box. The angle is taken in the unit
 *    square and the whole field is stretched to the box, so the isolines are
 *    the stretched perpendiculars (a 45 degree gradient on a wide box has its
 *    iso-lines parallel to the box's other diagonal). A CSS `linear-gradient`
 *    always keeps its isolines perpendicular to its angle, so no CSS angle
 *    reproduces this.
 *  - `path="circle"` whose `fillToRect` is not the box centre. The outer
 *    boundary is the circle circumscribing the box (radius = half the
 *    diagonal) and stop 0 sits at the focus, with every ray from the focus
 *    running to that circle: an SVG radial gradient with a focal point.
 *
 * Callers pair the result with `background-size: 100% 100%` and
 * `background-repeat: no-repeat`.
 *
 * @module gradient-svg-image
 */
import { escapeSvgAttr } from './visual-effects';

/** A gradient stop on the 0-100 position scale. */
export interface SvgImageGradientStop {
	color: string;
	position: number;
	opacity?: number;
}

function num(value: number): string {
	return String(Math.round(value * 10000) / 10000);
}

function stopMarkup(stops: readonly SvgImageGradientStop[]): string {
	return stops
		.map(
			(stop) =>
				`<stop offset="${num(Math.min(Math.max(stop.position, 0), 100) / 100)}" stop-color="${escapeSvgAttr(
					stop.color,
				)}"${typeof stop.opacity === 'number' ? ` stop-opacity="${num(stop.opacity)}"` : ''}/>`,
		)
		.join('');
}

function toCssUrl(svg: string): string {
	return `url("data:image/svg+xml,${encodeURIComponent(svg).replace(/[()]/gu, (c) => (c === '(' ? '%28' : '%29'))}")`;
}

/**
 * A linear gradient in `objectBoundingBox` space: the angle is measured in the
 * unit square and stretched to the box, with the first and last stops on the
 * box corners the angle reaches (`|cos| + |sin|` span), as PowerPoint draws a
 * `scaled="1"` gradient.
 */
export function buildScaledLinearGradientImage(
	stops: readonly SvgImageGradientStop[],
	ooxmlAngleDegrees: number,
): string {
	const radians = (ooxmlAngleDegrees * Math.PI) / 180;
	const dx = Math.cos(radians);
	const dy = Math.sin(radians);
	const span = (Math.abs(dx) + Math.abs(dy)) / 2;
	const svg =
		`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1" preserveAspectRatio="none">` +
		`<defs><linearGradient id="g" x1="${num(0.5 - dx * span)}" y1="${num(0.5 - dy * span)}" x2="${num(
			0.5 + dx * span,
		)}" y2="${num(0.5 + dy * span)}">${stopMarkup(stops)}</linearGradient></defs>` +
		`<rect width="1" height="1" fill="url(#g)"/></svg>`;
	return toCssUrl(svg);
}

/**
 * A `path="circle"` gradient whose focus (`fx`, `fy`, fractions of the box) is
 * off-centre: a true circle (in the box's own pixels, so `width` and `height`
 * are needed) circumscribing the box, with the focus as SVG's focal point.
 */
export function buildFocalCircleGradientImage(
	stops: readonly SvgImageGradientStop[],
	fx: number,
	fy: number,
	width: number,
	height: number,
): string {
	const radius = Math.hypot(width, height) / 2;
	const svg =
		`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${num(width)} ${num(height)}" preserveAspectRatio="none">` +
		`<defs><radialGradient id="g" gradientUnits="userSpaceOnUse" cx="${num(width / 2)}" cy="${num(
			height / 2,
		)}" r="${num(radius)}" fx="${num(fx * width)}" fy="${num(fy * height)}">${stopMarkup(stops)}</radialGradient></defs>` +
		`<rect width="${num(width)}" height="${num(height)}" fill="url(#g)"/></svg>`;
	return toCssUrl(svg);
}
