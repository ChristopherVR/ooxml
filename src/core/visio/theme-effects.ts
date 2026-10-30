/**
 * Theme effect styles. A shape whose shadow, glow, soft edge or reflection cells are `Themed`
 * takes them from the `a:effectLst` of the effect style its Quick Style effects matrix selects in
 * the effect scheme (EffectSchemeIndex) or, for a connector, the connector scheme. Only the effects
 * the renderer draws are read: `a:outerShdw`, `a:glow`, `a:softEdge` and `a:reflection`.
 */
import type { VisioGlow, VisioReflection, VisioShadow } from './model';
import type { Cells } from './sheet';
import { colorChoice, drawingPaint, themeChild, themeChildren } from './theme-color';
import { themeFormat, type ThemeResources } from './theme-resolve';

export interface ThemeEffects {
	shadow?: VisioShadow;
	glow?: VisioGlow;
	/** Soft edge radius in inches. */
	softEdges?: number;
	reflection?: VisioReflection;
}
const EMU = 914_400;
const clamp = (value: number, low: number, high: number) => Math.max(low, Math.min(high, value));
const measure = (node: Element, name: string, fallback = 0): number => {
	const raw = node.getAttribute(name);
	return raw !== null && /^-?\d{1,10}$/.test(raw) ? Number(raw) : fallback;
};
const paint = (node: Element) => drawingPaint(colorChoice(node), new Map());

/** The effects of the selected theme effect style, or undefined when none is selected. */
export function themeEffects(cells: Cells, resources: ThemeResources): ThemeEffects | undefined {
	const style = themeFormat(cells, 'Effects', resources);
	const list = themeChild(style, 'effectLst');
	if (!list) return undefined;
	const result: ThemeEffects = {};
	const shadow = themeChildren(list, 'outerShdw')[0];
	const shadowPaint = shadow && paint(shadow);
	if (shadow && shadowPaint) {
		const distance = clamp(measure(shadow, 'dist'), 0, 100 * EMU) / EMU;
		const angle = (measure(shadow, 'dir') / 60_000) * (Math.PI / 180);
		result.shadow = {
			color: shadowPaint.color,
			opacity: shadowPaint.opacity,
			offsetX: distance * Math.cos(angle),
			offsetY: -distance * Math.sin(angle),
			blur: clamp(measure(shadow, 'blurRad'), 0, 10 * EMU) / EMU,
		};
	}
	const glow = themeChildren(list, 'glow')[0];
	const glowPaint = glow && paint(glow);
	const radius = glow ? clamp(measure(glow, 'rad'), 0, 2.5 * EMU) / EMU : 0;
	if (glowPaint && radius > 0)
		result.glow = { color: glowPaint.color, opacity: glowPaint.opacity, size: radius };
	const soft = themeChildren(list, 'softEdge')[0];
	const softRadius = soft ? clamp(measure(soft, 'rad'), 0, 2 * EMU) / EMU : 0;
	if (softRadius > 0) result.softEdges = softRadius;
	const reflection = themeChildren(list, 'reflection')[0];
	if (reflection) {
		const opacity = clamp(measure(reflection, 'stA', 100_000), 0, 100_000) / 100_000;
		const size = clamp(measure(reflection, 'endPos', 100_000), 0, 100_000) / 100_000;
		if (opacity > 0 && size > 0)
			result.reflection = {
				opacity,
				size,
				distance: clamp(measure(reflection, 'dist'), 0, 2 * EMU) / EMU,
				blur: clamp(measure(reflection, 'blurRad'), 0, 2 * EMU) / EMU,
			};
	}
	return result;
}
