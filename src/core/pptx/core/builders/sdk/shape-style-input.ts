/**
 * Map the SDK's fill, stroke and effect inputs onto the {@link ShapeStyle}
 * model fields the save path already writes.
 *
 * @module sdk/shape-style-input
 */

import type { ConnectorArrowType } from '../../types/common';
import type { PptxCustomDashSegment, ShapeStyle } from '../../types/shape-style';
import type { GlowInput, InnerShadowInput, ReflectionInput } from './shape-style-input-types';
import type { FillInput, ShadowInput, StrokeInput } from './types';

/** `a:custDash/a:ds` lengths are 1000ths of a percent of the line width. */
const DASH_UNITS_PER_WIDTH = 100000;
/** `a:miter/@lim` is 1000ths of a percent of the line width. */
const MITER_UNITS_PER_RATIO = 100000;

export function mapFillToShapeStyle(fill?: FillInput): Partial<ShapeStyle> {
	if (!fill) {
		return {};
	}
	switch (fill.type) {
		case 'solid':
			return {
				fillMode: 'solid',
				fillColor: fill.color,
				fillOpacity: fill.opacity,
				...(fill.themeColorRef ? { fillColorRef: fill.themeColorRef } : {}),
			};
		case 'gradient':
			return {
				fillMode: 'gradient',
				fillGradientType: fill.path ? 'radial' : (fill.gradientType ?? 'linear'),
				fillGradientAngle: fill.angle ?? 0,
				fillGradientStops: fill.stops.map((s) => ({
					color: s.color,
					position: s.position,
					opacity: s.opacity,
				})),
				...(fill.path ? { fillGradientPathType: fill.path } : {}),
				...(fill.focus ? { fillGradientFillToRect: { ...fill.focus } } : {}),
				...(fill.tile ? { fillGradientTileRect: { ...fill.tile } } : {}),
				...(fill.scaled !== undefined ? { fillGradientScaled: fill.scaled } : {}),
				...(fill.rotWithShape !== undefined ? { fillGradientRotWithShape: fill.rotWithShape } : {}),
			};
		case 'pattern':
			return {
				fillMode: 'pattern',
				fillPatternPreset: fill.preset,
				...(fill.foreground !== undefined ? { fillColor: fill.foreground } : {}),
				fillPatternBackgroundColor: fill.background,
			};
		case 'image':
			return {
				fillMode: 'image',
				fillImageUrl: fill.url,
				fillImageMode: fill.mode ?? 'stretch',
			};
		case 'none':
			return { fillMode: 'none' };
		default:
			return {};
	}
}

/**
 * Turn px dash and gap lengths into `a:ds` segments relative to the line
 * width. An odd-length list repeats once, as SVG's `stroke-dasharray` does.
 */
function customDashSegments(lengths: number[], width: number): PptxCustomDashSegment[] {
	const values = lengths.length % 2 === 1 ? [...lengths, ...lengths] : lengths;
	const segments: PptxCustomDashSegment[] = [];
	for (let i = 0; i + 1 < values.length; i += 2) {
		segments.push({
			dash: Math.max(0, Math.round((values[i] / width) * DASH_UNITS_PER_WIDTH)),
			space: Math.max(0, Math.round((values[i + 1] / width) * DASH_UNITS_PER_WIDTH)),
		});
	}
	return segments;
}

function mapStrokeDash(stroke: StrokeInput): Partial<ShapeStyle> {
	if (!Array.isArray(stroke.dash)) {
		return { strokeDash: stroke.dash };
	}
	const width = stroke.width !== undefined && stroke.width > 0 ? stroke.width : 1;
	return {
		strokeDash: 'custom',
		customDashSegments: customDashSegments(stroke.dash, width),
	};
}

export function mapStrokeToShapeStyle(stroke?: StrokeInput): Partial<ShapeStyle> {
	if (!stroke) {
		return {};
	}
	const lineJoin = stroke.join ?? (stroke.miterLimit !== undefined ? 'miter' : undefined);
	// Only `a:miter` carries a limit (CT_LineJoinMiterProperties); with an
	// explicit round or bevel join there is nowhere to write one, so the model
	// does not claim it either.
	const miterLimit = lineJoin === 'miter' ? stroke.miterLimit : undefined;
	return {
		strokeColor: stroke.color,
		strokeWidth: stroke.width,
		...mapStrokeDash(stroke),
		strokeOpacity: stroke.opacity,
		lineJoin,
		lineCap: stroke.cap,
		...(stroke.themeColorRef ? { strokeColorRef: stroke.themeColorRef } : {}),
		...(miterLimit !== undefined
			? { miterLimit: Math.round(miterLimit * MITER_UNITS_PER_RATIO) }
			: {}),
		...(stroke.compound ? { compoundLine: stroke.compound } : {}),
		...(stroke.align ? { lineAlignment: stroke.align } : {}),
		...(stroke.headEnd
			? {
					connectorStartArrow: stroke.headEnd.type,
					connectorStartArrowWidth: stroke.headEnd.width,
					connectorStartArrowLength: stroke.headEnd.length,
				}
			: {}),
		...(stroke.tailEnd
			? {
					connectorEndArrow: stroke.tailEnd.type,
					connectorEndArrowWidth: stroke.tailEnd.width,
					connectorEndArrowLength: stroke.tailEnd.length,
				}
			: {}),
	};
}

/**
 * A connector's explicit `startArrow` / `endArrow` replaces that end of the
 * stroke whole: its size comes back to the default rather than carrying over
 * from a `stroke.headEnd` / `stroke.tailEnd` of a different type.
 */
export function connectorLineEndOverride(
	end: 'Start' | 'End',
	type: ConnectorArrowType | undefined,
): Partial<ShapeStyle> {
	if (type === undefined) {
		return {};
	}
	return {
		[`connector${end}Arrow`]: type,
		[`connector${end}ArrowWidth`]: undefined,
		[`connector${end}ArrowLength`]: undefined,
	};
}

function mapShadowToShapeStyle(shadow?: ShadowInput): Partial<ShapeStyle> {
	if (!shadow) {
		return {};
	}
	return {
		shadowColor: shadow.color ?? '#000000',
		shadowBlur: shadow.blur ?? 4,
		shadowOffsetX: shadow.offsetX ?? 2,
		shadowOffsetY: shadow.offsetY ?? 2,
		shadowOpacity: shadow.opacity ?? 0.4,
	};
}

function mapInnerShadowToShapeStyle(shadow?: InnerShadowInput): Partial<ShapeStyle> {
	if (!shadow) {
		return {};
	}
	return {
		innerShadowColor: shadow.color ?? '#000000',
		innerShadowBlur: shadow.blur,
		innerShadowOffsetX: shadow.offsetX,
		innerShadowOffsetY: shadow.offsetY,
		innerShadowOpacity: shadow.opacity,
	};
}

function mapGlowToShapeStyle(glow?: GlowInput): Partial<ShapeStyle> {
	if (!glow) {
		return {};
	}
	// `GlowInput.opacity` documents a fully opaque default; resolve it here
	// rather than fall through to the glow writer's own 0.4.
	return { glowColor: glow.color, glowRadius: glow.radius, glowOpacity: glow.opacity ?? 1 };
}

function mapReflectionToShapeStyle(reflection?: ReflectionInput): Partial<ShapeStyle> {
	if (!reflection) {
		return {};
	}
	return {
		reflectionBlurRadius: reflection.blur,
		reflectionDistance: reflection.distance,
		reflectionStartOpacity: reflection.startOpacity ?? 0.5,
		reflectionEndOpacity: reflection.endOpacity ?? 0,
		reflectionEndPosition: reflection.endPosition ?? 0.9,
		// PowerPoint's reflection presets: a mirror image below the shape.
		reflectionDirection: 90,
		reflectionScaleY: -100000,
		reflectionAlignment: 'bl',
		reflectionRotateWithShape: false,
	};
}

/** The fill, stroke and effect inputs a shape-like element accepts. */
export interface ShapeStyleInputs {
	fill?: FillInput;
	stroke?: StrokeInput;
	shadow?: ShadowInput;
	innerShadow?: InnerShadowInput;
	glow?: GlowInput;
	softEdge?: number;
	reflection?: ReflectionInput;
}

export function buildShapeStyle(opts: ShapeStyleInputs): ShapeStyle | undefined {
	const parts: Partial<ShapeStyle> = {
		...mapFillToShapeStyle(opts.fill),
		...mapStrokeToShapeStyle(opts.stroke),
		...mapShadowToShapeStyle(opts.shadow),
		...mapInnerShadowToShapeStyle(opts.innerShadow),
		...mapGlowToShapeStyle(opts.glow),
		...(opts.softEdge !== undefined ? { softEdgeRadius: opts.softEdge } : {}),
		...mapReflectionToShapeStyle(opts.reflection),
	};
	return Object.keys(parts).length > 0 ? (parts as ShapeStyle) : undefined;
}
