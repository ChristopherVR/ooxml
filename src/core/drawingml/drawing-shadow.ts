import { EMU_PER_PIXEL } from '../units/constants';
import { parseXml, type XmlElement } from '../xml/index';
import { resolveDrawingColor, type DrawingColorTheme } from './drawing-color';
import { parseDrawingEffects } from './drawing-effects';
import type { DrawingOuterShadow } from './effect-model';

/** Resolved pixel geometry shared by SVG effect painters. */
export interface DrawingSvgShadow {
	color: string;
	opacity: number;
	blur: number;
	dx: number;
	dy: number;
}

/** Ordinary outer shadows only. Projected scale/skew shadows require an affine painter. */
export function resolveDrawingShadow(
	effects: XmlElement | undefined,
	theme: DrawingColorTheme,
): DrawingSvgShadow | undefined {
	const effect = parseDrawingEffects(effects)?.effects.find(
		(entry) => entry.kind === 'outerShadow',
	);
	return effect?.kind === 'outerShadow' ? resolveDrawingOuterShadow(effect, theme) : undefined;
}

/** Projects the shared authored shadow model into the existing SVG painter contract. */
export function resolveDrawingOuterShadow(
	effect: DrawingOuterShadow,
	theme: DrawingColorTheme,
): DrawingSvgShadow | undefined {
	// Explicit malformed source geometry must not turn into a default zero shadow.
	if (effect.sourceXml) {
		const source = parseXml(effect.sourceXml).documentElement;
		for (const attribute of ['blurRad', 'dist', 'dir', 'sx', 'sy', 'kx', 'ky']) {
			const value = source.getAttribute(attribute);
			if (value !== null && !Number.isFinite(Number(value))) return undefined;
		}
	}
	if (
		(effect.scaleX ?? 100000) !== 100000 ||
		(effect.scaleY ?? 100000) !== 100000 ||
		(effect.skewXAngle ?? 0) !== 0 ||
		(effect.skewYAngle ?? 0) !== 0
	)
		return undefined;
	const color = effect.color;
	const resolved = color && resolveDrawingColor(color, theme, { transformOrder: 'document' });
	const blur = (effect.blurRadiusEmu ?? 0) / EMU_PER_PIXEL;
	const distance = (effect.distanceEmu ?? 0) / EMU_PER_PIXEL;
	const angle = (((effect.directionAngle ?? 0) / 60000) * Math.PI) / 180;
	if (!resolved || ![blur, distance, angle].every(Number.isFinite) || blur < 0 || distance < 0)
		return undefined;
	return {
		color: resolved.hex,
		opacity: resolved.alpha,
		blur,
		dx: Math.cos(angle) * distance,
		dy: Math.sin(angle) * distance,
	};
}

export function resolveDrawingShadowXml(
	xml: string | undefined,
	theme: DrawingColorTheme,
): DrawingSvgShadow | undefined {
	return xml ? resolveDrawingShadow(parseXml(xml).documentElement, theme) : undefined;
}

/** Extracted from PowerPoint's gallery effect painter; callers provide existing formatters. */
export function svgDropShadowElement(
	s: DrawingSvgShadow,
	format: { number(value: number): string; color(value: string, fallback: string): string },
): string {
	const num = format.number;
	return `<feDropShadow dx="${num(s.dx)}" dy="${num(s.dy)}" stdDeviation="${num(s.blur / 2)}" flood-color="${format.color(s.color, '#000000')}" flood-opacity="${num(s.opacity)}" result="shadow"/>`;
}
