import { themeColorRefFromColorChoice } from '../../color/theme-color-ref';
import { TextStyle, XmlObject } from '../../types';
import { extractColorChoiceXml } from '../../utils/color-xml-preservation';
import { buildEffectDagTreeFromXml } from '../builders/effect-dag-containers';
import { extractReflectionAttributes } from '../builders/effect-style-extractor-reflection';
import { findObjectEffect, shadowOffset } from '../builders/effect-style-from-model';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeTextStyleUtils';
import { isFiniteNumber } from '../../../../units/finite';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Parse hyperlink-related properties (a:hlinkClick, a:hlinkMouseOver) from
	 * run-level XML and apply them to the given TextStyle.
	 */
	protected applyHyperlinkStyle(
		style: TextStyle,
		runProperties: XmlObject,
		relationshipMap?: Map<string, string>,
	): void {
		const hyperlinkNode = runProperties['a:hlinkClick'] as XmlObject | undefined;
		const hyperlinkRelationshipId = String(
			hyperlinkNode?.['@_r:id'] || hyperlinkNode?.['@_id'] || '',
		).trim();
		if (hyperlinkRelationshipId.length > 0) {
			style.hyperlinkRId = hyperlinkRelationshipId;
			const hyperlinkTarget = relationshipMap?.get(hyperlinkRelationshipId);
			if (hyperlinkTarget) {
				style.hyperlink = hyperlinkTarget;
			}
		}
		const tooltip = String(hyperlinkNode?.['@_tooltip'] || '').trim();
		if (tooltip) {
			style.hyperlinkTooltip = tooltip;
		}
		// Additional hyperlink attributes
		if (hyperlinkNode) {
			const invalidUrl = String(hyperlinkNode['@_invalidUrl'] || '').trim();
			if (invalidUrl) {
				style.hyperlinkInvalidUrl = invalidUrl;
			}
			const tgtFrame = String(hyperlinkNode['@_tgtFrame'] || '').trim();
			if (tgtFrame) {
				style.hyperlinkTargetFrame = tgtFrame;
			}
			const historyAttr = hyperlinkNode['@_history'];
			if (historyAttr !== undefined) {
				const hVal = String(historyAttr).trim().toLowerCase();
				style.hyperlinkHistory = hVal !== '0' && hVal !== 'false';
			}
			const highlightClick = hyperlinkNode['@_highlightClick'];
			if (highlightClick !== undefined) {
				const hcVal = String(highlightClick).trim().toLowerCase();
				style.hyperlinkHighlightClick = hcVal === '1' || hcVal === 'true';
			}
			const endSnd = hyperlinkNode['@_endSnd'];
			if (endSnd !== undefined) {
				const esVal = String(endSnd).trim().toLowerCase();
				style.hyperlinkEndSound = esVal === '1' || esVal === 'true';
			}
			// Preserve the embedded-WAV `a:snd` child (r:embed + name) verbatim so
			// the hyperlink click sound survives a round-trip.
			const clickSnd = hyperlinkNode['a:snd'];
			if (clickSnd && typeof clickSnd === 'object') {
				style.hyperlinkSoundXml = clickSnd as XmlObject;
			}
			// Preserve `a:extLst` verbatim (most commonly Microsoft's
			// `ahyp:hlinkClr` "hyperlink colour" extension) so an unmodelled
			// vendor extension does not vanish on save.
			const extLst = hyperlinkNode['a:extLst'];
			if (extLst && typeof extLst === 'object') {
				style.hyperlinkExtensionXml = extLst as XmlObject;
			}
		}
		const actionStr = String(hyperlinkNode?.['@_action'] || '').trim();
		if (actionStr) {
			style.hyperlinkAction = actionStr;
			// For action hyperlinks without an explicit URL (ppaction:// verbs),
			// derive a user-visible hyperlink so the link is clickable.
			if (!style.hyperlink && actionStr.startsWith('ppaction://')) {
				style.hyperlink = actionStr;
			}
		}
		// Resolve internal slide jump targets from action hyperlinks
		if (
			actionStr === 'ppaction://hlinksldjump' &&
			hyperlinkRelationshipId.length > 0 &&
			relationshipMap
		) {
			const slideTarget = relationshipMap.get(hyperlinkRelationshipId);
			if (slideTarget) {
				const slideMatch = slideTarget.match(/slide(\d+)\.xml$/i);
				if (slideMatch) {
					// The file number is NOT the slide's position: decks that were
					// reordered or had slides deleted keep their part names. Resolve
					// through the presentation's slide order (as shape-level actions
					// do) and fall back to the file number only when the part is not
					// in the list (e.g. while parsing a layout before slides load).
					const slideNumber = parseInt(slideMatch[1] ?? '', 10);
					const orderIndex = this.orderedSlidePaths.findIndex(
						(slidePath) => slidePath.match(/slide(\d+)\.xml$/i)?.[1] === String(slideNumber),
					);
					style.hyperlinkTargetSlideIndex = orderIndex >= 0 ? orderIndex : slideNumber - 1;
				}
			}
		}

		const hlinkMouseOver = runProperties['a:hlinkMouseOver'] as XmlObject | undefined;
		if (hlinkMouseOver) {
			const mouseOverRelId = String(
				hlinkMouseOver['@_r:id'] || hlinkMouseOver['@_id'] || '',
			).trim();
			if (mouseOverRelId.length > 0) {
				const mouseOverTarget = relationshipMap?.get(mouseOverRelId);
				if (mouseOverTarget) {
					style.hyperlinkMouseOver = mouseOverTarget;
				} else {
					style.hyperlinkMouseOver = mouseOverRelId;
				}
				// Preserve the mouse-over `a:snd` child verbatim for round-trip.
				const mouseOverSnd = hlinkMouseOver['a:snd'];
				if (mouseOverSnd && typeof mouseOverSnd === 'object') {
					style.hyperlinkMouseOverSoundXml = mouseOverSnd as XmlObject;
				}
			}
		}
	}

	/**
	 * Parse text-run-level effects (a:effectLst on a:rPr): outer shadow, glow,
	 * and reflection — and apply them to the given TextStyle.
	 */
	protected applyTextRunEffects(style: TextStyle, runEffectList: XmlObject): void {
		// The CT_EffectList members are read by the shared DrawingML effect reader
		// (`findObjectEffect`, over the object-tree view); colours and opacity
		// resolve through the runtime colour codec on the effect's node, so
		// scheme colours follow the slide's theme and colour map.
		const emuPerPx = PptxHandlerRuntimeBase.EMU_PER_PX;
		const px = (emu: number | undefined) =>
			emu !== undefined && Math.trunc(emu) >= 0 ? Math.trunc(emu) / emuPerPx : undefined;
		const degrees = (angle: number | undefined) =>
			angle === undefined ? undefined : Math.trunc(angle) / 60000;

		// Outer shadow
		const outer = findObjectEffect(runEffectList, 'outerShadow');
		if (outer) {
			const shdwColor = this.parseColor(outer.node);
			if (shdwColor) {
				style.textShadowColor = shdwColor;
			}
			// No `a:alpha` is an opaque shadow (ECMA-376), not the renderer's 0.5.
			style.textShadowOpacity = this.extractColorOpacity(outer.node) ?? 1;
			const blur = px(outer.effect.blurRadiusEmu);
			if (blur !== undefined) {
				style.textShadowBlur = blur;
			}
			const offset = shadowOffset(
				px(outer.effect.distanceEmu),
				degrees(outer.effect.directionAngle) ?? 0,
			);
			if (offset.x !== undefined && offset.y !== undefined) {
				style.textShadowOffsetX = offset.x;
				style.textShadowOffsetY = offset.y;
			}
		}
		// Glow
		const glow = findObjectEffect(runEffectList, 'glow');
		if (glow) {
			const glowColor = this.parseColor(glow.node);
			if (glowColor) {
				style.textGlowColor = glowColor;
			}
			// Preserve the original colour choice (e.g. `a:schemeClr`) so the
			// writer can re-emit it instead of always resolving to a flat
			// `a:srgbClr`, which would sever the glow from theme/Recolor changes.
			const glowColorXml = extractColorChoiceXml(glow.node);
			if (glowColorXml) {
				style.textGlowColorXml = glowColorXml;
			}
			const glowColorRef = themeColorRefFromColorChoice(glow.node);
			if (glowColorRef) {
				style.textGlowColorRef = glowColorRef;
			}
			const glowOpacity = this.extractColorOpacity(glow.node);
			if (glowOpacity !== undefined) {
				style.textGlowOpacity = glowOpacity;
			}
			const radius = px(glow.effect.radiusEmu);
			if (radius !== undefined) {
				style.textGlowRadius = radius;
			}
		}
		// Reflection. Reuses the same `extractReflectionAttributes` helper the
		// shape-level parser (`PptxShapeEffectStyleExtractor`) relies on, over
		// the shared effect reader, so a text run's `a:reflection` honours the
		// full attribute set.
		const reflection = findObjectEffect(runEffectList, 'reflection');
		if (reflection) {
			style.textReflection = true;
			const attrs = extractReflectionAttributes(reflection.node, emuPerPx);
			const fields = [
				['reflectionBlurRadius', 'textReflectionBlur'],
				['reflectionStartOpacity', 'textReflectionStartOpacity'],
				['reflectionEndOpacity', 'textReflectionEndOpacity'],
				['reflectionDistance', 'textReflectionOffset'],
				['reflectionEndPosition', 'textReflectionEndPosition'],
				['reflectionDirection', 'textReflectionDirection'],
				['reflectionFadeDirection', 'textReflectionFadeDirection'],
				['reflectionScaleX', 'textReflectionScaleX'],
				['reflectionScaleY', 'textReflectionScaleY'],
				['reflectionSkewX', 'textReflectionSkewX'],
				['reflectionSkewY', 'textReflectionSkewY'],
				['reflectionRotation', 'textReflectionRotation'],
				['reflectionAlignment', 'textReflectionAlignment'],
			] as const;
			for (const [from, to] of fields) {
				if (attrs[from] !== undefined) {
					(style as Record<string, unknown>)[to] = attrs[from];
				}
			}
		}

		// Inner shadow
		const inner = findObjectEffect(runEffectList, 'innerShadow');
		if (inner) {
			const color = this.parseColor(inner.node);
			if (color) {
				style.textInnerShadowColor = color;
			}
			// Preserve the original colour choice (e.g. `a:prstClr`) so the
			// writer can re-emit it instead of always resolving to a flat
			// `a:srgbClr`, which would silently swap an authored preset colour
			// for its resolved hex on every round trip.
			const innerShadowColorXml = extractColorChoiceXml(inner.node);
			if (innerShadowColorXml) {
				style.textInnerShadowColorXml = innerShadowColorXml;
			}
			const innerShadowColorRef = themeColorRefFromColorChoice(inner.node);
			if (innerShadowColorRef) {
				style.textInnerShadowColorRef = innerShadowColorRef;
			}
			const opacity = this.extractColorOpacity(inner.node);
			if (opacity !== undefined) {
				style.textInnerShadowOpacity = opacity;
			}
			const blur = px(inner.effect.blurRadiusEmu);
			if (blur !== undefined) {
				style.textInnerShadowBlur = blur;
			}
			const offset = shadowOffset(
				px(inner.effect.distanceEmu),
				degrees(inner.effect.directionAngle) ?? 0,
			);
			if (offset.x !== undefined && offset.y !== undefined) {
				style.textInnerShadowOffsetX = offset.x;
				style.textInnerShadowOffsetY = offset.y;
			}
		}

		// Preset shadow
		const preset = findObjectEffect(runEffectList, 'presetShadow');
		if (preset) {
			const prst = preset.effect.preset?.trim();
			if (prst) {
				style.textPresetShadowName = prst;
			}
			const color = this.parseColor(preset.node);
			if (color) {
				style.textPresetShadowColor = color;
			}
			const opacity = this.extractColorOpacity(preset.node);
			if (opacity !== undefined) {
				style.textPresetShadowOpacity = opacity;
			}
			const distance = px(preset.effect.distanceEmu);
			if (distance !== undefined) {
				style.textPresetShadowDistance = distance;
			}
			const direction = degrees(preset.effect.directionAngle);
			if (direction !== undefined) {
				style.textPresetShadowDirection = direction;
			}
		}

		// Blur
		const blur = px(findObjectEffect(runEffectList, 'blur')?.effect.radiusEmu);
		if (blur !== undefined) {
			style.textBlurRadius = blur;
		}

		// Soft edge (feathers the glyph's own alpha silhouette; see the
		// `textSoftEdgeRadius` doc comment for why this never rendered before).
		const softEdge = px(findObjectEffect(runEffectList, 'softEdge')?.effect.radiusEmu);
		if (softEdge !== undefined) {
			style.textSoftEdgeRadius = softEdge;
		}

		// The members below are not CT_EffectList children (they belong to
		// blip effects), but some producers put them on a run's `a:effectLst`
		// and PowerPoint's reader has always kept them.
		// Alpha modification fixed
		const alphaModFix = runEffectList['a:alphaModFix'] as XmlObject | undefined;
		if (alphaModFix) {
			const amt = Number.parseInt(String(alphaModFix['@_amt'] || ''), 10);
			if (isFiniteNumber(amt)) {
				style.textAlphaModFix = amt / 1000;
			}
		}

		// Alpha modulation
		const alphaMod = runEffectList['a:alphaMod'] as XmlObject | undefined;
		if (alphaMod) {
			const amt = Number.parseInt(String(alphaMod['@_amt'] || ''), 10);
			if (isFiniteNumber(amt)) {
				style.textAlphaMod = amt / 1000;
			}
		}

		// Hue/saturation/luminance
		const hslNode = runEffectList['a:hsl'] as XmlObject | undefined;
		if (hslNode) {
			const hue = Number.parseInt(String(hslNode['@_hue'] || ''), 10);
			if (isFiniteNumber(hue)) {
				style.textHslHue = hue / 60000;
			}
			const sat = Number.parseInt(String(hslNode['@_sat'] || ''), 10);
			if (isFiniteNumber(sat)) {
				style.textHslSaturation = sat / 1000;
			}
			const lum = Number.parseInt(String(hslNode['@_lum'] || ''), 10);
			if (isFiniteNumber(lum)) {
				style.textHslLuminance = lum / 1000;
			}
		}

		// Colour change
		const clrChange = runEffectList['a:clrChange'] as XmlObject | undefined;
		if (clrChange) {
			const from = this.parseColor(clrChange['a:clrFrom'] as XmlObject | undefined);
			const to = this.parseColor(clrChange['a:clrTo'] as XmlObject | undefined);
			if (from) {
				style.textClrChangeFrom = from;
			}
			if (to) {
				style.textClrChangeTo = to;
			}
		}

		// Duotone
		const duotoneNode = runEffectList['a:duotone'] as XmlObject | undefined;
		if (duotoneNode) {
			const colorNodes = [
				...this.ensureArray(duotoneNode['a:srgbClr']),
				...this.ensureArray(duotoneNode['a:schemeClr']),
				...this.ensureArray(duotoneNode['a:prstClr']),
			] as XmlObject[];
			if (colorNodes.length >= 2) {
				style.textDuotone = {
					color1: this.parseColor(colorNodes[0]) || '#000000',
					color2: this.parseColor(colorNodes[1]) || '#ffffff',
				};
			}
		}
	}

	/**
	 * Parse `a:effectDag` on `a:rPr` (run-level effect graph) and attach both
	 * the raw XML (for verbatim round-trip) and a typed structural tree to
	 * the given TextStyle. ECMA-376 §21.1.2.3.6 lists `a:effectDag` as a
	 * valid alternative to `a:effectLst` inside CT_TextCharacterProperties;
	 * historically the codebase ignored it on the run side, so any preset
	 * carrying a run-level DAG would lose its effects on save.
	 *
	 * The four structural container nodes (`a:cont`, `a:blend`,
	 * `a:xfrmEffect`, `a:relOff`) are parsed into a typed
	 * {@link import('../../types').EffectDagContainer} tree; everything else
	 * (e.g. `a:outerShdw`, `a:glow`, `a:alphaInv`) is captured opaquely as
	 * raw XML so we don't duplicate the entire effect taxonomy here.
	 */
	protected applyTextRunEffectDag(style: TextStyle, runProperties: XmlObject): void {
		const effectDag = runProperties['a:effectDag'] as XmlObject | undefined;
		if (!effectDag) {
			return;
		}
		style.textEffectDagXml = effectDag;
		const tree = buildEffectDagTreeFromXml(effectDag);
		if (tree) {
			style.textEffectDagTree = tree;
		}
	}
}
