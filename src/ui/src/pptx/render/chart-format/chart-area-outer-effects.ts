/**
 * The chart area's outer shadow and glow, which PowerPoint draws OUTSIDE the
 * chart frame (the chart area is the whole frame, so an effect drawn inside
 * the chart's own SVG would be clipped at its edges).
 *
 * They are expressed as the element-level effect style every binding already
 * applies to an element's container (`getComputedEffectStyle` turns it into a
 * `box-shadow` and a glow `drop-shadow` filter), so all five bindings draw
 * them with no binding change.
 *
 * @module render/chart-format/chart-area-outer-effects
 */
import type { PptxChartData, ShapeStyle } from 'ooxml-core/pptx';

/** Slide px per point: the slide coordinate space is 96 dpi. */
const PX_PER_PT = 96 / 72;

/**
 * The effect style of a chart's frame from its chart-area format, or
 * `undefined` when the chart area casts nothing. A chart area with no fill
 * and no border has no shape to cast from, so it casts nothing either.
 */
export function chartAreaEffectStyle(data: PptxChartData | undefined): ShapeStyle | undefined {
	const format = data?.elementFormats?.chartArea;
	const effects = format && !format.reset ? format.effects : undefined;
	if (!effects || (!effects.shadow && !(effects.glow && effects.glow.radius > 0))) {
		return undefined;
	}
	const fill = format?.fill?.kind ?? (data?.style?.chartAreaFill === 'none' ? 'none' : 'automatic');
	const line = format?.line?.kind ?? (data?.style?.chartAreaBorder ? 'solid' : 'automatic');
	if (fill === 'none' && line !== 'solid' && line !== 'gradient') {
		return undefined;
	}
	const style: ShapeStyle = {};
	if (effects.shadow) {
		const s = effects.shadow;
		style.shadowColor = s.color;
		style.shadowOpacity = 1 - (s.transparency ?? 0) / 100;
		style.shadowBlur = s.blur * PX_PER_PT;
		style.shadowDistance = s.distance * PX_PER_PT;
		style.shadowAngle = s.angle;
		// Authored DrawingML semantics: `blurRad` is a radius (halved for CSS).
		style.outerShadowXml = {};
	}
	if (effects.glow && effects.glow.radius > 0) {
		style.glowColor = effects.glow.color;
		style.glowRadius = effects.glow.radius * PX_PER_PT;
		style.glowOpacity = 1 - (effects.glow.transparency ?? 0) / 100;
	}
	return style;
}
