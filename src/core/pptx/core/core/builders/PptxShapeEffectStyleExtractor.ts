import type { ShapeStyle, XmlObject } from '../../types';
import { effectChild } from './effect-list-roundtrip';
import { extractFillOverlayAttributes } from './effect-style-extractor-fill-overlay';
import {
	blurStyle,
	findObjectEffect,
	glowStyle,
	innerShadowStyle,
	outerShadowStyle,
	reflectionStyle,
	softEdgeStyle,
} from './effect-style-from-model';
import { PRESET_SHADOW_BLUR_MAP, PRESET_SHADOW_OPACITY_MAP } from './effect-style-preset-maps';

export interface PptxShapeEffectStyleExtractorContext {
	emuPerPx: number;
	parseColor: (colorNode: XmlObject | undefined, placeholderColor?: string) => string | undefined;
	extractColorOpacity: (colorNode: XmlObject | undefined) => number | undefined;
	ensureArray: (value: unknown) => XmlObject[];
}

export interface IPptxShapeEffectStyleExtractor {
	extractShadowStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
	extractPresetShadowStyle(effectLstParent: XmlObject): Partial<ShapeStyle>;
	extractInnerShadowStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
	extractGlowStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
	extractSoftEdgeStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
	extractReflectionStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
	extractBlurStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
	extractFillOverlayStyle(shapeProps: XmlObject): Partial<ShapeStyle>;
}

export class PptxShapeEffectStyleExtractor implements IPptxShapeEffectStyleExtractor {
	private readonly context: PptxShapeEffectStyleExtractorContext;

	public constructor(context: PptxShapeEffectStyleExtractorContext) {
		this.context = context;
	}

	public extractShadowStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		const outer = findObjectEffect(effectChild(shapeProps, 'effectLst'), 'outerShadow');
		if (!outer) {
			return this.extractPresetShadowStyle(shapeProps);
		}
		return outerShadowStyle(outer.effect, outer.node, this.context);
	}

	public extractPresetShadowStyle(effectLstParent: XmlObject): Partial<ShapeStyle> {
		const preset = findObjectEffect(
			effectLstParent['a:effectLst'] as XmlObject | undefined,
			'presetShadow',
		);
		if (!preset) {
			return {};
		}
		const { effect, node } = preset;
		const name = String(effect.preset ?? '').trim();
		const distPx = (effect.distanceEmu ?? 0) / this.context.emuPerPx;
		const dirRad = ((effect.directionAngle ?? 0) / 60000) * (Math.PI / 180);

		// Look up preset-specific blur values (shdw1-shdw20)
		const presetBlur = PRESET_SHADOW_BLUR_MAP[name];
		const parsedOpacity = this.context.extractColorOpacity(node);

		return {
			shadowOffsetX: Math.round(distPx * Math.cos(dirRad) * 100) / 100,
			shadowOffsetY: Math.round(distPx * Math.sin(dirRad) * 100) / 100,
			shadowColor: this.context.parseColor(node) || '#000000',
			shadowOpacity:
				parsedOpacity ??
				(presetBlur !== undefined ? (PRESET_SHADOW_OPACITY_MAP[name] ?? 0.5) : 0.5),
			shadowBlur: presetBlur ?? 4,
			presetShadowName: name.length > 0 ? name : undefined,
		};
	}

	public extractInnerShadowStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		const inner = findObjectEffect(effectChild(shapeProps, 'effectLst'), 'innerShadow');
		return inner ? innerShadowStyle(inner.effect, inner.node, this.context) : {};
	}

	public extractGlowStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		const glow = findObjectEffect(effectChild(shapeProps, 'effectLst'), 'glow');
		return glow ? glowStyle(glow.effect, glow.node, this.context) : {};
	}

	public extractSoftEdgeStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		const softEdge = findObjectEffect(effectChild(shapeProps, 'effectLst'), 'softEdge');
		return softEdge ? softEdgeStyle(softEdge.effect, this.context) : {};
	}

	public extractReflectionStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		const reflection = findObjectEffect(effectChild(shapeProps, 'effectLst'), 'reflection');
		return reflection
			? reflectionStyle(reflection.effect, reflection.node, this.context.emuPerPx)
			: {};
	}

	public extractBlurStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		const blur = findObjectEffect(shapeProps['a:effectLst'] as XmlObject | undefined, 'blur');
		return blur ? blurStyle(blur.effect, this.context) : {};
	}

	/**
	 * `a:fillOverlay` as a DIRECT child of the plain `a:effectLst`; see
	 * {@link extractFillOverlayAttributes} for the full explanation (split into
	 * its own module to keep this file under the size limit).
	 */
	public extractFillOverlayStyle(shapeProps: XmlObject): Partial<ShapeStyle> {
		return extractFillOverlayAttributes(shapeProps, this.context);
	}
}

export { PRESET_SHADOW_BLUR_MAP, PRESET_SHADOW_OPACITY_MAP } from './effect-style-preset-maps';
