import {
	gradientFocusCenter,
	readDrawingGradient,
	sortGradientStopEntries,
	type DrawingGradientRead,
} from '../../../../drawingml/gradient-read';
import { OBJECT_DRAWING_VIEW } from '../../color/object-drawing-view';
import { themeColorRefFromColorChoice, themeColorRefToXml } from '../../color/theme-color-ref';
import type { ShapeStyle, XmlObject } from '../../types';
import { ooxmlGradientAngleToCssDegrees } from '../../utils/gradient-angle';
import { drawingChild, mergeDrawingFillXml } from './drawing-fill-xml';

/**
 * Parse `a:gradFill/a:tileRect` LTRB insets into fractions (1/100000 units, may be negative,
 * an unreadable side reads 0). Standalone so the shape-style extractor can use it directly.
 */
export function extractGradientTileRect(
	gradFill: XmlObject,
): NonNullable<ShapeStyle['fillGradientTileRect']> | undefined {
	return finiteSides(readObjectGradient(gradFill).tileRectSides);
}

/** `a:gradFill` read by the shared `drawingml` reader over the object tree. */
export function readObjectGradient(gradFill: XmlObject): DrawingGradientRead<XmlObject> {
	return readDrawingGradient(OBJECT_DRAWING_VIEW, gradFill);
}

/** LTRB sides with each unreadable side as 0, as the codec always read them. */
function finiteSides(
	sides: { l: number; t: number; r: number; b: number } | undefined,
): { l: number; t: number; r: number; b: number } | undefined {
	if (!sides) {
		return undefined;
	}
	const side = (value: number) => (Number.isFinite(value) ? value : 0);
	return { l: side(sides.l), t: side(sides.t), r: side(sides.r), b: side(sides.b) };
}

/**
 * Re-emit an authored `a:fillToRect` / `a:tileRect` verbatim when its
 * resolved LTRB fractions still match what was parsed from it.
 *
 * `CT_RelativeRect`'s `l`/`t`/`r`/`b` attributes are all independently
 * optional, each defaulting to 0 when absent, so a source that authored only
 * some of the four (e.g. `r`/`b` for a bottom-right focal point, leaving
 * `l`/`t` to default) must not gain the other two purely because the writer
 * always fills in all four when it has anything to write at all. A genuine
 * edit (the resolved value no longer matches what parse captured) still gets
 * a freshly built, fully explicit node, matching how every other numeric
 * override in this codebase falls back to an explicit re-serialisation once
 * it no longer agrees with the source.
 */
function resolveRelativeRectXml(
	original: XmlObject | undefined,
	current: { l: number; t: number; r: number; b: number },
	originalResolved: { l: number; t: number; r: number; b: number } | undefined,
): XmlObject {
	if (
		original &&
		originalResolved &&
		originalResolved.l === current.l &&
		originalResolved.t === current.t &&
		originalResolved.r === current.r &&
		originalResolved.b === current.b
	) {
		return original;
	}
	return {
		'@_l': String(Math.round(current.l * 100000)),
		'@_t': String(Math.round(current.t * 100000)),
		'@_r': String(Math.round(current.r * 100000)),
		'@_b': String(Math.round(current.b * 100000)),
	};
}

export interface PptxGradientStyleCodecContext {
	ensureArray: (value: unknown) => unknown[];
	parseColor: (colorNode: XmlObject | undefined, placeholderColor?: string) => string | undefined;
	extractColorOpacity: (colorNode: XmlObject | undefined) => number | undefined;
	clampUnitInterval: (value: number) => number;
	hexToRgb: (hex: string) => { r: number; g: number; b: number } | undefined;
	rgbToHex: (r: number, g: number, b: number) => string;
}

export interface IPptxGradientStyleCodec {
	extractGradientOpacity(gradFill: XmlObject): number | undefined;
	extractGradientStops(gradFill: XmlObject): NonNullable<ShapeStyle['fillGradientStops']>;
	extractGradientType(gradFill: XmlObject): NonNullable<ShapeStyle['fillGradientType']>;
	extractGradientAngle(gradFill: XmlObject): number;
	extractGradientPathType(gradFill: XmlObject): ShapeStyle['fillGradientPathType'];
	extractGradientFocalPoint(gradFill: XmlObject): ShapeStyle['fillGradientFocalPoint'];
	extractGradientFillToRect(gradFill: XmlObject): ShapeStyle['fillGradientFillToRect'];
	extractGradientTileRect(gradFill: XmlObject): ShapeStyle['fillGradientTileRect'];
	extractGradientFlip(gradFill: XmlObject): ShapeStyle['fillGradientFlip'];
	extractGradientRotWithShape(gradFill: XmlObject): boolean | undefined;
	extractGradientScaled(gradFill: XmlObject): boolean | undefined;
	buildGradientCssFromStops(
		stops: NonNullable<ShapeStyle['fillGradientStops']>,
		type: NonNullable<ShapeStyle['fillGradientType']>,
		angle: number,
		focalPoint?: ShapeStyle['fillGradientFocalPoint'],
		pathType?: ShapeStyle['fillGradientPathType'],
		fillToRect?: ShapeStyle['fillGradientFillToRect'],
	): string | undefined;
	extractGradientFillCss(gradFill: XmlObject): string | undefined;
	buildGradientFillXml(shapeStyle: ShapeStyle): XmlObject | undefined;
	extractGradientFillColor(gradFill: XmlObject): string | undefined;
}

export class PptxGradientStyleCodec implements IPptxGradientStyleCodec {
	private readonly context: PptxGradientStyleCodecContext;

	public constructor(context: PptxGradientStyleCodecContext) {
		this.context = context;
	}

	public extractGradientOpacity(gradFill: XmlObject): number | undefined {
		const opacities = readObjectGradient(gradFill)
			.stops.map((stop) => this.context.extractColorOpacity(stop.node))
			.filter((opacity): opacity is number => opacity !== undefined);
		if (opacities.length === 0) {
			return undefined;
		}

		const opacityTotal = opacities.reduce((sum, value) => sum + value, 0);
		return this.context.clampUnitInterval(opacityTotal / opacities.length);
	}

	public extractGradientStops(gradFill: XmlObject): NonNullable<ShapeStyle['fillGradientStops']> {
		// A stop without a readable `pos` (required by the schema) sits at 0, as it always did.
		const sortedStops = sortGradientStopEntries(readObjectGradient(gradFill).stops, 0);
		const stops: NonNullable<ShapeStyle['fillGradientStops']> = [];
		for (const { node: stop, position: rawPosition } of sortedStops) {
			const color = this.context.parseColor(stop);
			if (!color) {
				continue;
			}

			const position = this.context.clampUnitInterval((rawPosition ?? 0) / 100000) * 100;
			const opacity = this.context.extractColorOpacity(stop);

			// Preserve original color XML for round-trip (scheme colors + transforms)
			const originalColorXml = this.extractColorSubNode(stop);
			// `stop` (the `a:gs` node) carries its colour choice as a direct
			// child, the same shape `themeColorRefFromColorChoice` expects from
			// an `a:solidFill` wrapper, so it can be passed straight through.
			const colorRef = themeColorRefFromColorChoice(stop);

			stops.push({
				color,
				position,
				opacity,
				...(originalColorXml ? { originalColorXml } : {}),
				...(colorRef ? { colorRef } : {}),
			});
		}
		return stops;
	}

	public extractGradientType(gradFill: XmlObject): NonNullable<ShapeStyle['fillGradientType']> {
		return readObjectGradient(gradFill).hasPath ? 'radial' : 'linear';
	}

	public extractGradientPathType(gradFill: XmlObject): ShapeStyle['fillGradientPathType'] {
		const gradient = readObjectGradient(gradFill);
		if (!gradient.hasPath) {
			return undefined;
		}
		const pathValue = String(gradient.fill.path ?? '')
			.trim()
			.toLowerCase();
		if (pathValue === 'circle' || pathValue === 'rect' || pathValue === 'shape') {
			return pathValue;
		}
		// Default to "circle" when a:path exists but path attr is missing
		return 'circle';
	}

	public extractGradientFocalPoint(gradFill: XmlObject): ShapeStyle['fillGradientFocalPoint'] {
		// The focal point is the centre of the rectangle from (l, t) to (1 - r, 1 - b).
		const sides = finiteSides(readObjectGradient(gradFill).fillToRectSides);
		return sides ? gradientFocusCenter(sides) : undefined;
	}

	public extractGradientFillToRect(gradFill: XmlObject): ShapeStyle['fillGradientFillToRect'] {
		const sides = finiteSides(readObjectGradient(gradFill).fillToRectSides);
		if (!sides) {
			return undefined;
		}
		const clamp = (value: number) => this.context.clampUnitInterval(value);
		return { l: clamp(sides.l), t: clamp(sides.t), r: clamp(sides.r), b: clamp(sides.b) };
	}

	public extractGradientTileRect(gradFill: XmlObject): ShapeStyle['fillGradientTileRect'] {
		return extractGradientTileRect(gradFill);
	}

	public extractGradientFlip(gradFill: XmlObject): ShapeStyle['fillGradientFlip'] {
		const flip = readObjectGradient(gradFill).flip;
		if (flip === 'x' || flip === 'y' || flip === 'xy' || flip === 'none') {
			return flip;
		}
		return undefined;
	}

	public extractGradientRotWithShape(gradFill: XmlObject): boolean | undefined {
		return readObjectGradient(gradFill).rotateWithShape;
	}

	public extractGradientScaled(gradFill: XmlObject): boolean | undefined {
		return readObjectGradient(gradFill).fill.scaled;
	}

	public extractGradientAngle(gradFill: XmlObject): number {
		const angle = readObjectGradient(gradFill).fill.angle;
		return angle === undefined ? 90 : ((angle % 360) + 360) % 360;
	}

	public buildGradientCssFromStops(
		stops: NonNullable<ShapeStyle['fillGradientStops']>,
		type: NonNullable<ShapeStyle['fillGradientType']>,
		angle: number,
		focalPoint?: ShapeStyle['fillGradientFocalPoint'],
		pathType?: ShapeStyle['fillGradientPathType'],
		fillToRect?: ShapeStyle['fillGradientFillToRect'],
	): string | undefined {
		if (stops.length === 0) {
			return undefined;
		}

		const stopTokens = stops.map((stop) => {
			const colorToken = this.colorWithOpacity(stop.color, stop.opacity);
			const normalizedPosition =
				typeof stop.position === 'number' && Number.isFinite(stop.position)
					? `${this.context.clampUnitInterval(stop.position / 100) * 100}%`
					: undefined;
			return normalizedPosition ? `${colorToken} ${normalizedPosition}` : colorToken;
		});

		if (type === 'radial') {
			const stopStr = stopTokens.join(', ');
			const resolvedPathType = pathType || 'circle';

			// Compute center from fillToRect if available
			let cx = 50;
			let cy = 50;
			if (fillToRect) {
				cx = ((fillToRect.l + (1 - fillToRect.r)) / 2) * 100;
				cy = ((fillToRect.t + (1 - fillToRect.b)) / 2) * 100;
			} else if (focalPoint) {
				cx = focalPoint.x * 100;
				cy = focalPoint.y * 100;
			}
			const posX = `${Math.round(cx)}%`;
			const posY = `${Math.round(cy)}%`;

			if (resolvedPathType === 'rect') {
				const semiX = Math.max(cx, 100 - cx);
				const semiY = Math.max(cy, 100 - cy);
				return `radial-gradient(${Math.round(semiX)}% ${Math.round(semiY)}% at ${posX} ${posY}, ${stopStr})`;
			}

			if (resolvedPathType === 'shape') {
				return `radial-gradient(farthest-side at ${posX} ${posY}, ${stopStr})`;
			}

			return `radial-gradient(circle at ${posX} ${posY}, ${stopStr})`;
		}
		// `angle` is an OOXML `a:lin/@ang` value in degrees (clockwise from +x);
		// CSS measures clockwise from "to top", a quarter turn away.
		return `linear-gradient(${ooxmlGradientAngleToCssDegrees(angle).toFixed(2)}deg, ${stopTokens.join(', ')})`;
	}

	public extractGradientFillCss(gradFill: XmlObject): string | undefined {
		const stops = this.extractGradientStops(gradFill);
		if (stops.length === 0) {
			return undefined;
		}

		return this.buildGradientCssFromStops(
			stops,
			this.extractGradientType(gradFill),
			this.extractGradientAngle(gradFill),
			this.extractGradientFocalPoint(gradFill),
			this.extractGradientPathType(gradFill),
			this.extractGradientFillToRect(gradFill),
		);
	}

	public buildGradientFillXml(shapeStyle: ShapeStyle): XmlObject | undefined {
		const stops = (shapeStyle.fillGradientStops || [])
			.filter((stop) => Boolean(stop?.color))
			.map((stop) => {
				const positionRaw =
					typeof stop.position === 'number' && Number.isFinite(stop.position) ? stop.position : 0;
				const position = Math.round(this.context.clampUnitInterval(positionRaw / 100) * 100000);

				// A typed theme ref wins over everything else: it keeps this stop
				// following the theme palette after a later theme change, which a
				// re-emitted `a:srgbClr` (or even a preserved `a:schemeClr` that
				// happens to still resolve to the same hex) cannot express.
				if (stop.colorRef) {
					const refStopXml: XmlObject = {
						'@_pos': String(position),
						'a:schemeClr': themeColorRefToXml(stop.colorRef),
					};
					return refStopXml;
				}

				// Prefer original color XML to preserve scheme colors and transforms
				if (
					stop.originalColorXml &&
					this.context.parseColor(stop.originalColorXml) === String(stop.color).trim()
				) {
					const stopXml: XmlObject = {
						'@_pos': String(position),
						...stop.originalColorXml,
					};
					return stopXml;
				}

				const normalizedColor = String(stop.color || '').trim();
				const normalizedOpacity =
					typeof stop.opacity === 'number' && Number.isFinite(stop.opacity)
						? this.context.clampUnitInterval(stop.opacity)
						: undefined;
				const stopXml: XmlObject = {
					'@_pos': String(position),
					'a:srgbClr': {
						'@_val': normalizedColor.replace('#', ''),
					},
				};
				if (normalizedOpacity !== undefined) {
					(stopXml['a:srgbClr'] as XmlObject)['a:alpha'] = {
						'@_val': String(Math.round(normalizedOpacity * 100000)),
					};
				}
				return stopXml;
			});
		if (stops.length === 0) {
			return undefined;
		}

		const gradientType = shapeStyle.fillGradientType || 'linear';
		const gradientXml: XmlObject = {};
		// `@flip` and `@rotWithShape` attach to the `<a:gradFill>` element
		// and must be set before the child `<a:gsLst>` is inserted —
		// fast-xml-parser preserves insertion order, but attributes always
		// serialise on the parent regardless. We assign them up-front for
		// readability.
		if (shapeStyle.fillGradientFlip) {
			// `extractGradientFlip` only ever returns a value (including
			// `'none'`) when the source explicitly authored `@flip`: `'none'`
			// is ALSO ECMA-376's schema default for `ST_TileFlipMode`, but a
			// value the source explicitly wrote is not ours to drop just
			// because PowerPoint would have assumed the same thing anyway.
			// `gradientXml` is a freshly built object (not merged onto a
			// preserved node), so setting the key to `undefined` here would
			// not "remove" anything; it would just never write it, which is
			// exactly the bug: an authored `flip="none"` came back with no
			// `@flip` attribute at all.
			gradientXml['@_flip'] = shapeStyle.fillGradientFlip;
		}
		if (shapeStyle.fillGradientRotWithShape !== undefined) {
			gradientXml['@_rotWithShape'] = shapeStyle.fillGradientRotWithShape ? '1' : '0';
		}
		gradientXml['a:gsLst'] = { 'a:gs': stops };
		if (gradientType === 'radial') {
			const pathType = shapeStyle.fillGradientPathType || 'circle';
			const pathXml: XmlObject = {
				'@_path': pathType,
			};
			if (shapeStyle.fillGradientFillToRect) {
				const ftr = shapeStyle.fillGradientFillToRect;
				const originalPath = drawingChild(shapeStyle.fillGradientXml, 'path');
				const originalFillToRect = drawingChild(originalPath, 'fillToRect');
				const originalResolved = shapeStyle.fillGradientXml
					? this.extractGradientFillToRect(shapeStyle.fillGradientXml)
					: undefined;
				pathXml['a:fillToRect'] = resolveRelativeRectXml(originalFillToRect, ftr, originalResolved);
			} else if (shapeStyle.fillGradientFocalPoint) {
				const fp = shapeStyle.fillGradientFocalPoint;
				// Convert focal point back to fillToRect LTRB values
				const l = Math.round(fp.x * 100000);
				const t = Math.round(fp.y * 100000);
				const r = Math.round((1 - fp.x) * 100000);
				const b = Math.round((1 - fp.y) * 100000);
				pathXml['a:fillToRect'] = {
					'@_l': String(l),
					'@_t': String(t),
					'@_r': String(r),
					'@_b': String(b),
				};
			}
			gradientXml['a:path'] = pathXml;
		} else {
			const normalizedAngle =
				typeof shapeStyle.fillGradientAngle === 'number' &&
				Number.isFinite(shapeStyle.fillGradientAngle)
					? shapeStyle.fillGradientAngle
					: 90;
			const linNode: XmlObject = {
				'@_ang': String(Math.round(normalizedAngle * 60000)),
			};
			// `@scaled` defaults to "1" per the schema, but we only emit
			// the explicit value the parser captured to avoid lying about
			// what the source authored.
			if (shapeStyle.fillGradientScaled !== undefined) {
				linNode['@_scaled'] = shapeStyle.fillGradientScaled ? '1' : '0';
			} else {
				linNode['@_scaled'] = '1';
			}
			gradientXml['a:lin'] = linNode;
		}
		if (shapeStyle.fillGradientTileRect) {
			const tr = shapeStyle.fillGradientTileRect;
			const originalTileRect = drawingChild(shapeStyle.fillGradientXml, 'tileRect');
			const originalResolved = shapeStyle.fillGradientXml
				? extractGradientTileRect(shapeStyle.fillGradientXml)
				: undefined;
			gradientXml['a:tileRect'] = resolveRelativeRectXml(originalTileRect, tr, originalResolved);
		}
		return mergeDrawingFillXml(
			shapeStyle.fillGradientXml,
			gradientXml,
			['gsLst', 'lin', 'path'],
			['gsLst', 'lin', 'path', 'tileRect', 'extLst'],
		);
	}

	public extractGradientFillColor(gradFill: XmlObject): string | undefined {
		const sortedStops = this.extractStopsSorted(gradFill);
		if (sortedStops.length === 0) {
			return undefined;
		}

		const firstColor = this.context.parseColor(sortedStops[0]);
		const lastColor =
			sortedStops.length > 1
				? this.context.parseColor(sortedStops[sortedStops.length - 1])
				: undefined;

		if (!firstColor) {
			return lastColor;
		}
		if (!lastColor) {
			return firstColor;
		}

		const firstRgb = this.context.hexToRgb(firstColor);
		const lastRgb = this.context.hexToRgb(lastColor);
		if (!firstRgb || !lastRgb) {
			return firstColor;
		}

		return this.context.rgbToHex(
			(firstRgb.r + lastRgb.r) / 2,
			(firstRgb.g + lastRgb.g) / 2,
			(firstRgb.b + lastRgb.b) / 2,
		);
	}

	private extractStopsSorted(gradFill: XmlObject): XmlObject[] {
		return sortGradientStopEntries(readObjectGradient(gradFill).stops, 0).map((stop) => stop.node);
	}

	/**
	 * Extract the raw color sub-node from a gradient stop for round-trip preservation.
	 * Returns the color choice node (e.g. a:schemeClr with transforms, a:srgbClr, etc.).
	 */
	private extractColorSubNode(stop: XmlObject): XmlObject | undefined {
		const colorNames = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);
		for (const [key, value] of Object.entries(stop)) {
			if (colorNames.has(key.split(':').at(-1) ?? key)) {
				return { [key]: value } as XmlObject;
			}
		}
		return undefined;
	}

	private colorWithOpacity(color: string, opacity: number | undefined): string {
		if (opacity === undefined) {
			return color;
		}
		const rgb = this.context.hexToRgb(color);
		if (!rgb) {
			return color;
		}
		return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${this.context.clampUnitInterval(opacity)})`;
	}
}
