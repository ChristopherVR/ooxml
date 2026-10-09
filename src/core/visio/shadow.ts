import type { VisioShadow } from './model';
import { number, type Cells, type Report } from './sheet';

/**
 * Saved shape shadow (Fill Format section): ShdwPattern, ShdwForegnd, ShdwForegndTrans and the
 * ShapeShdw offset and blur cells. Only a simple, unscaled outer shadow is modelled.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/shapeshdwtype-cell-fill-format-section
 */
export function shapeShadow(
	cells: Cells,
	color: (name: string, fallback: string) => string,
	report: Report,
): VisioShadow | undefined {
	if (cells.get('ShdwPattern')?.value === 'Themed') {
		report('unsupported-theme-shadow', 'Theme-selected shadows are not rendered.');
		return undefined;
	}
	const pattern = number(cells, 'ShdwPattern', 0, report);
	if (!pattern) return undefined;
	if (pattern !== 1)
		report('unsupported-shadow-pattern', `Shadow pattern ${pattern} is drawn as a solid shadow.`);
	const type = number(cells, 'ShapeShdwType', 0, report);
	if (type === 3)
		report('unsupported-oblique-shadow', 'Oblique shadows are drawn as offset shadows.');
	const scale = number(cells, 'ShapeShdwScaleFactor', 1, report);
	if (Math.abs(scale - 1) > 1e-9)
		report('unsupported-shadow-scale', 'Scaled shadows are drawn at the shape size.');
	const hasOffset = cells.has('ShapeShdwOffsetX') && cells.has('ShapeShdwOffsetY');
	if (!hasOffset)
		report(
			'approximate-shadow-offset',
			"A shadow without a shape offset uses Visio's default 0.125 inch page offset.",
		);
	const clamp = (value: number, limit: number) => Math.max(-limit, Math.min(limit, value));
	const transparency = number(cells, 'ShdwForegndTrans', 0, report);
	report(
		'unverified-shadow',
		"Shadows are drawn as an offset, blurred copy and may differ from Microsoft Visio's rendering.",
	);
	return {
		color: color('ShdwForegnd', '#000000'),
		opacity: Math.max(0, Math.min(1, 1 - transparency)),
		offsetX: clamp(hasOffset ? number(cells, 'ShapeShdwOffsetX', 0.125, report) : 0.125, 1000),
		offsetY: clamp(hasOffset ? number(cells, 'ShapeShdwOffsetY', -0.125, report) : -0.125, 1000),
		blur: Math.max(0, Math.min(10, number(cells, 'ShapeShdwBlur', 0, report))),
	};
}
