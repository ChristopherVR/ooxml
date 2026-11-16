import {
	interpretDrawingEffectDag,
	readDrawingEffectContainer,
	type DrawingDagColorOp,
} from '../../../../drawingml/effect-dag';
import type { DrawingEffectFields } from '../../../../drawingml/effect-read';
import { OBJECT_DRAWING_VIEW } from '../../color/object-drawing-view';
import type { ShapeStyle, XmlObject } from '../../types';
import { buildEffectDagTreeFromXml } from './effect-dag-containers';
import { extractDagDuotone } from './effect-dag-specific-helpers';
import {
	blurStyle,
	glowStyle,
	innerShadowStyle,
	outerShadowStyle,
	reflectionStyle,
	softEdgeStyle,
} from './effect-style-from-model';

/**
 * Context required by the DAG extractor — a subset of the effect codec context.
 */
export interface PptxEffectDagExtractorContext {
	emuPerPx: number;
	parseColor: (colorNode: XmlObject | undefined, placeholderColor?: string) => string | undefined;
	extractColorOpacity: (colorNode: XmlObject | undefined) => number | undefined;
	ensureArray: (value: unknown) => XmlObject[];
}

export interface IPptxEffectDagExtractor {
	extractEffectDagStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
}

const VALID_FILL_OVERLAY_BLENDS = new Set(['over', 'mult', 'screen', 'darken', 'lighten']);

/**
 * Parses `a:effectDag`, the DAG alternative to `a:effectLst`, through the shared reader and
 * interpretation (`drawingml/effect-dag.ts`): the graph is walked from the shape along its paint
 * path, so effects inside `a:cont` and painted `a:blend` containers are found too. The first
 * effect of each list kind maps onto the same ShapeStyle fields as `a:effectLst`
 * (`effect-style-from-model.ts`); the colour operations (grayscale, biLevel, lum, hsl, tint,
 * duotone) and the combined alpha map onto the `dag*` fields. What the interpretation
 * approximates is listed in `effectDagApproximations`. The raw XML and the typed structural
 * tree are kept for round-trip saving.
 */
export class PptxEffectDagExtractor implements IPptxEffectDagExtractor {
	private readonly context: PptxEffectDagExtractorContext;

	public constructor(context: PptxEffectDagExtractorContext) {
		this.context = context;
	}

	/**
	 * Extract style properties from an `a:effectDag` element on `shapeProps`.
	 * Returns an empty object if no DAG is present.
	 */
	public extractEffectDagStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		const effectDag = childByLocalName(shapeProps, 'effectDag');
		if (!effectDag || typeof effectDag !== 'object') {
			return {};
		}

		const style: Partial<ShapeStyle> = { effectDagXml: effectDag };
		const tree = buildEffectDagTreeFromXml(effectDag);
		if (tree) {
			style.effectDagTree = tree;
		}

		const paint = interpretDrawingEffectDag(
			readDrawingEffectContainer(OBJECT_DRAWING_VIEW, effectDag)!,
		);
		const seen = new Set<string>();
		for (const { effect, source } of paint.effects) {
			if (seen.has(effect.kind) || !source) {
				continue;
			}
			seen.add(effect.kind);
			Object.assign(style, this.effectStyle(effect, source));
		}
		for (const op of paint.colorOps) {
			this.applyColorOp(op, style);
		}
		if (paint.opacity < 1) {
			style.dagAlphaModFix = paint.opacity * 100;
		}
		if (paint.approximations.length > 0) {
			style.effectDagApproximations = paint.approximations;
		}
		return style;
	}

	private effectStyle(effect: DrawingEffectFields, node: XmlObject): Partial<ShapeStyle> {
		const { emuPerPx } = this.context;
		switch (effect.kind) {
			case 'outerShadow':
				return outerShadowStyle(effect, node, this.context);
			case 'innerShadow':
				return innerShadowStyle(effect, node, this.context);
			case 'glow':
				return glowStyle(effect, node, this.context);
			case 'softEdge':
				return softEdgeStyle(effect, this.context);
			case 'reflection':
				return reflectionStyle(effect, node, emuPerPx);
			case 'blur':
				return blurStyle(effect, this.context);
			case 'fillOverlay':
				return this.fillOverlayStyle(effect.blend, node);
			default:
				// A preset shadow in a DAG keeps its source XML only, as before.
				return {};
		}
	}

	/**
	 * The `a:fillOverlay` blend mode and its overlay tint: the solid colour, or the first
	 * gradient stop's colour, so the renderer can paint the tint instead of only a blend mode.
	 */
	private fillOverlayStyle(blend: string | undefined, fillOverlay: XmlObject): Partial<ShapeStyle> {
		const style: Partial<ShapeStyle> = {};
		const mode = String(blend ?? '')
			.trim()
			.toLowerCase();
		if (VALID_FILL_OVERLAY_BLENDS.has(mode)) {
			style.dagFillOverlayBlend = mode as ShapeStyle['dagFillOverlayBlend'];
		}
		const colorHost = fillOverlayColorHost(fillOverlay, this.context);
		const color = colorHost ? this.context.parseColor(colorHost) : undefined;
		if (color) {
			style.dagFillOverlayColor = color;
			const opacity = this.context.extractColorOpacity(colorHost);
			if (typeof opacity === 'number' && Number.isFinite(opacity)) {
				style.dagFillOverlayOpacity = opacity;
			}
		}
		return style;
	}

	/** The first operation of each kind sets its `dag*` field (percentages as 0..100). */
	private applyColorOp(op: DrawingDagColorOp<XmlObject>, style: Partial<ShapeStyle>): void {
		switch (op.kind) {
			case 'grayscale':
				style.dagGrayscale ??= true;
				break;
			case 'biLevel':
				style.dagBiLevel ??= Math.max(0, Math.min(100, op.threshold * 100));
				break;
			case 'luminance':
				if (style.dagLumBrightness === undefined && style.dagLumContrast === undefined) {
					style.dagLumBrightness = op.brightness * 100;
					style.dagLumContrast = op.contrast * 100;
				}
				break;
			case 'hsl':
				if (style.dagHslHue === undefined) {
					style.dagHslHue = op.hue;
					// `@sat` adjusts saturation (0 = unchanged); the field is the resulting
					// percentage of the original that the painter's `saturate()` takes.
					style.dagHslSaturation = 100 + op.saturation * 100;
					style.dagHslLuminance = op.luminance * 100;
				}
				break;
			case 'tint':
				if (style.dagTintHue === undefined) {
					style.dagTintHue = op.hue;
					style.dagTintAmount = op.amount * 100;
				}
				break;
			case 'duotone':
				if (!style.dagDuotone && op.source) {
					extractDagDuotone({ 'a:duotone': op.source }, style, this.context);
				}
				break;
			default:
				break;
		}
	}
}

/**
 * Resolve the colour-bearing node inside an `a:fillOverlay`: the `a:solidFill`
 * directly, or the first gradient stop (`a:gradFill/a:gsLst/a:gs[0]`) for a
 * gradient overlay. Returns `undefined` when neither carries a colour.
 */
export function fillOverlayColorHost(
	fillOverlay: XmlObject,
	context: PptxEffectDagExtractorContext,
): XmlObject | undefined {
	const solidFill = fillOverlay['a:solidFill'] as XmlObject | undefined;
	if (solidFill) {
		return solidFill;
	}
	const gradFill = fillOverlay['a:gradFill'] as XmlObject | undefined;
	if (!gradFill) {
		return undefined;
	}
	const gsLst = gradFill['a:gsLst'] as XmlObject | undefined;
	if (!gsLst) {
		return undefined;
	}
	const stops = context.ensureArray(gsLst['a:gs']);
	return stops.length > 0 ? stops[0] : undefined;
}

function childByLocalName(xml: XmlObject, name: string): XmlObject | undefined {
	const entry = Object.entries(xml).find(([key]) => {
		const local = key.includes(':') ? key.slice(key.indexOf(':') + 1) : key;
		return !key.startsWith('@_') && local === name;
	});
	return entry?.[1] as XmlObject | undefined;
}

export type { DagSpecificContext } from './effect-dag-specific-helpers';
export { extractDagDuotone } from './effect-dag-specific-helpers';
export {
	parseEffectDagContainer,
	buildEffectDagTreeFromXml,
	serializeEffectDagContainer,
} from './effect-dag-containers';
