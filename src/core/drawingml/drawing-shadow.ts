import { EMU_PER_PIXEL } from '../units/constants';
import { NS, first, parseXml, type XmlElement } from '../xml/index';
import { parseDrawingColorIn, resolveDrawingColor, type DrawingColorTheme } from './drawing-color';

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
	const node = first(effects, 'outerShdw', NS.a);
	if (!node) return undefined;
	const value = (key: string, fallback: number) => Number(node.getAttribute(key) ?? fallback);
	if (
		value('sx', 100000) !== 100000 ||
		value('sy', 100000) !== 100000 ||
		value('kx', 0) !== 0 ||
		value('ky', 0) !== 0
	)
		return undefined;
	const color = parseDrawingColorIn(node);
	const resolved = color && resolveDrawingColor(color, theme, { transformOrder: 'document' });
	const blur = value('blurRad', 0) / EMU_PER_PIXEL;
	const distance = value('dist', 0) / EMU_PER_PIXEL;
	const angle = ((value('dir', 0) / 60000) * Math.PI) / 180;
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
