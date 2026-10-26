import type { ChartShapeProperties } from '../../../chart/index';
import { resolveDrawingColor } from '../../../drawingml/drawing-color';
import { parseDrawingEffects } from '../../../drawingml/drawing-effects';
import type { DrawingColor, DrawingFill } from '../../../drawingml/types';
import { parseXml } from '../../../xml/index';
import type {
	PptxChartElementEffects,
	PptxChartElementFill,
	PptxChartElementFormat,
	PptxChartElementLine,
	XmlObject,
} from '../types';
import { chartGradientFromNeutral } from './chart-from-neutral-paint';

const round = (value: number) => Math.round(value * 100) / 100;
type ParseColor = (node: XmlObject | undefined, placeholder?: string) => string | undefined;

/** Shared DrawingML fields projected only for compatibility rendering, never back into source. */
export function chartFormatFromNeutral(
	shape: ChartShapeProperties,
	parseColor: ParseColor,
	pictureFallback?: PptxChartElementFill,
): PptxChartElementFormat {
	const resolve = (color: DrawingColor | undefined) =>
		color
			? resolveDrawingColor(color, {
					scheme: (name) => parseColor({ 'a:schemeClr': { '@_val': name } }),
				})
			: undefined;
	const colorFields = (color: DrawingColor | undefined) => {
		const resolved = resolve(color);
		return {
			...(resolved ? { color: resolved.hex } : {}),
			...(resolved && resolved.alpha < 1
				? { transparency: round(100 * (1 - resolved.alpha)) }
				: {}),
		};
	};
	const paint = (fill: DrawingFill | undefined): PptxChartElementFill | undefined => {
		if (!fill) return undefined;
		if (fill.kind === 'none') return { kind: 'none' };
		if (fill.kind === 'solid') return { kind: 'solid', ...colorFields(fill.color) };
		if (fill.kind === 'gradient') {
			const gradient = chartGradientFromNeutral(fill, { parseColor });
			return gradient ? { kind: 'gradient', gradient } : undefined;
		}
		if (fill.kind === 'pattern')
			return {
				kind: 'pattern',
				pattern: {
					preset: fill.preset || 'pct5',
					foreground: resolve(fill.foreground)?.hex ?? '#000000',
					background: resolve(fill.background)?.hex ?? '#FFFFFF',
				},
			};
		return pictureFallback;
	};
	const format: PptxChartElementFormat = {};
	const fill = paint(shape.fill);
	if (fill) format.fill = fill;
	const source = shape.line;
	if (source) {
		const stroke = paint(source.fill);
		const line: PptxChartElementLine =
			stroke && ['none', 'solid', 'gradient'].includes(stroke.kind)
				? ({ ...stroke } as PptxChartElementLine)
				: { kind: 'automatic' };
		if (source.widthEmu !== undefined) line.width = round(source.widthEmu / 12700);
		if (['sng', 'dbl', 'thickThin', 'thinThick', 'tri'].includes(source.compound ?? ''))
			line.compound = source.compound as PptxChartElementLine['compound'];
		if (['rnd', 'sq', 'flat'].includes(source.cap ?? ''))
			line.cap = source.cap as PptxChartElementLine['cap'];
		if (source.dash) line.dash = source.dash;
		if (source.join) line.join = source.join.kind;
		for (const [field, end] of [
			['beginArrow', source.headEnd],
			['endArrow', source.tailEnd],
		] as const) {
			if (end?.type && ['triangle', 'stealth', 'diamond', 'oval', 'arrow'].includes(end.type))
				line[field] = end.type as NonNullable<PptxChartElementLine['beginArrow']>;
		}
		format.line = line;
	}
	if (shape.effectsXml) {
		const effects: PptxChartElementEffects = {};
		for (const effect of parseDrawingEffects(parseXml(shape.effectsXml).documentElement)?.effects ??
			[]) {
			if (effect.kind === 'outerShadow')
				effects.shadow ??= {
					color: resolve(effect.color)?.hex ?? '#000000',
					...colorFields(effect.color),
					blur: round((effect.blurRadiusEmu ?? 0) / 12700),
					distance: round((effect.distanceEmu ?? 0) / 12700),
					angle: round((effect.directionAngle ?? 0) / 60000),
				};
			if (effect.kind === 'glow')
				effects.glow ??= {
					color: resolve(effect.color)?.hex ?? '#FFFFFF',
					...colorFields(effect.color),
					radius: round((effect.radiusEmu ?? 0) / 12700),
				};
			if (effect.kind === 'softEdge') effects.softEdge ??= round((effect.radiusEmu ?? 0) / 12700);
		}
		if (Object.keys(effects).length) format.effects = effects;
	}
	return format;
}
