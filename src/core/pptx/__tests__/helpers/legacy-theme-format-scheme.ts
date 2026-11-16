/**
 * The pptx theme format scheme reader, style reference resolution and theme part writer as they
 * were before they moved onto the shared DrawingML theme model (docs/agnostic-core-plan.md step 3), frozen as the
 * reference the corpus parity harness compares the new reader with
 * (`integration/drawingml-theme-format-scheme-parity.test.ts`). Test support only: never import it
 * from shipped code.
 */
/* eslint-disable */
import type { PptxThemeColorScheme, ShapeStyle, XmlObject } from '../../core/types';
import type {
	PptxThemeEffectStyle,
	PptxThemeFillStyle,
	PptxThemeFormatScheme,
	PptxThemeLineStyle,
} from '../../core/types';
import { extractStyleReferenceColorXml, withThemePlaceholderColor } from '../../core/utils';
import { PptxHandlerRuntime } from '../../core/core/runtime/PptxHandlerRuntimeImplementation';

/**
 * Ordered fill-child tag tokens recognised inside `a:fillStyleLst` /
 * `a:bgFillStyleLst`. Used by {@link extractFillStyleListChildOrder} to
 * recover document order from raw XML when fast-xml-parser has merged
 * heterogeneous siblings into typed buckets.
 */
const FILL_LIST_CHILD_TAGS = [
	'a:solidFill',
	'a:gradFill',
	'a:pattFill',
	'a:noFill',
	'a:grpFill',
] as const;

type FillListChildTag = (typeof FILL_LIST_CHILD_TAGS)[number];

/**
 * Recover the document order of fill children in an `a:fillStyleLst` /
 * `a:bgFillStyleLst` block from a raw XML string by scanning the opening
 * tags between the list element's `<a:fillStyleLst[ \t\n>]` and its
 * closing tag. Returns an array of child tag names in source order.
 *
 * Returns an empty array when the listing tag isn't found or the raw XML
 * is unavailable.
 */
function extractFillStyleListChildOrder(
	rawXml: string | undefined,
	listTag: 'a:fillStyleLst' | 'a:bgFillStyleLst',
): FillListChildTag[] {
	if (!rawXml) {
		return [];
	}
	const openRegex = new RegExp(`<${listTag.replace(':', '\\:')}\\b[^>]*>`);
	const closeRegex = new RegExp(`</${listTag.replace(':', '\\:')}\\s*>`);
	const openMatch = openRegex.exec(rawXml);
	if (!openMatch) {
		return [];
	}
	const startIdx = openMatch.index + openMatch[0].length;
	const closeMatch = closeRegex.exec(rawXml.slice(startIdx));
	if (!closeMatch) {
		return [];
	}
	const inner = rawXml.slice(startIdx, startIdx + closeMatch.index);
	const order: FillListChildTag[] = [];
	const childRegex = /<a:(solidFill|gradFill|pattFill|noFill|grpFill)\b/g;
	let match: RegExpExecArray | null;
	while ((match = childRegex.exec(inner)) !== null) {
		order.push(`a:${match[1]}` as FillListChildTag);
	}
	return order;
}

export class LegacyThemeFormatSchemeRuntime extends PptxHandlerRuntime {
	/**
	 * Build a fill-style entry for one child node based on its tag.
	 */
	public legacyBuildFillStyleEntry(tag: FillListChildTag, node: XmlObject): PptxThemeFillStyle {
		switch (tag) {
			case 'a:solidFill':
				return {
					kind: 'solid',
					color: this.parseColor(node),
					opacity: this.extractColorOpacity(node),
					rawNode: node,
				};
			case 'a:gradFill':
				return {
					kind: 'gradient',
					color: this.extractGradientFillColor(node),
					opacity: this.extractGradientOpacity(node),
					gradientStops: this.extractGradientStops(node),
					gradientAngle: this.extractGradientAngle(node),
					gradientType: this.extractGradientType(node),
					gradientCss: this.extractGradientFillCss(node),
					rawNode: node,
				};
			case 'a:pattFill':
				return {
					kind: 'pattern',
					color:
						this.parseColor(node['a:fgClr'] as XmlObject | undefined) ||
						this.parseColor(node['a:bgClr'] as XmlObject | undefined),
					patternPreset: String(node['@_prst'] || '').trim() || undefined,
					patternBackgroundColor:
						this.parseColor(node['a:bgClr'] as XmlObject | undefined) || undefined,
					rawNode: node,
				};
			case 'a:noFill':
				return { kind: 'none', rawNode: node };
			case 'a:grpFill':
				return { kind: 'group', rawNode: node };
		}
	}

	/**
	 * Collect fill-style children from a style list node, preserving
	 * document order when an `orderHint` is supplied (extracted from the
	 * raw theme XML). Handles `a:solidFill`, `a:gradFill`, `a:pattFill`,
	 * `a:noFill`, `a:grpFill`.
	 *
	 * fast-xml-parser without `preserveOrder` collapses heterogeneous
	 * siblings into per-tag buckets and we lose the relative document
	 * order. When the caller supplies an `orderHint` (a list of tag names
	 * in source order recovered via {@link extractFillStyleListChildOrder})
	 * we walk the buckets in step with the hint and emit entries in true
	 * document order. Without a hint we fall back to the historical
	 * solid → gradient → pattern → noFill → grpFill grouping.
	 */
	public legacyCollectFillChildren(
		listNode: XmlObject,
		orderHint?: readonly FillListChildTag[],
	): PptxThemeFillStyle[] {
		const buckets: Record<FillListChildTag, XmlObject[]> = {
			'a:solidFill': this.ensureArray(listNode['a:solidFill']) as XmlObject[],
			'a:gradFill': this.ensureArray(listNode['a:gradFill']) as XmlObject[],
			'a:pattFill': this.ensureArray(listNode['a:pattFill']) as XmlObject[],
			'a:noFill': this.ensureArray(listNode['a:noFill']) as XmlObject[],
			'a:grpFill': this.ensureArray(listNode['a:grpFill']) as XmlObject[],
		};

		if (orderHint && orderHint.length > 0) {
			const cursors: Record<FillListChildTag, number> = {
				'a:solidFill': 0,
				'a:gradFill': 0,
				'a:pattFill': 0,
				'a:noFill': 0,
				'a:grpFill': 0,
			};
			const ordered: PptxThemeFillStyle[] = [];
			for (const tag of orderHint) {
				const idx = cursors[tag];
				const node = buckets[tag][idx];
				if (node) {
					ordered.push(this.legacyBuildFillStyleEntry(tag, node));
					cursors[tag] = idx + 1;
				}
			}
			// Append any nodes the hint did not cover (defensive).
			for (const tag of FILL_LIST_CHILD_TAGS) {
				while (cursors[tag] < buckets[tag].length) {
					ordered.push(this.legacyBuildFillStyleEntry(tag, buckets[tag][cursors[tag]]));
					cursors[tag] += 1;
				}
			}
			return ordered;
		}

		// No order hint, fall back to typed grouping (legacy behaviour).
		const results: PptxThemeFillStyle[] = [];
		for (const tag of FILL_LIST_CHILD_TAGS) {
			for (const node of buckets[tag]) {
				results.push(this.legacyBuildFillStyleEntry(tag, node));
			}
		}
		return results;
	}

	/**
	 * Parse each child of a `a:fillStyleLst` (or `a:bgFillStyleLst`).
	 * Children can be `a:solidFill`, `a:gradFill`, `a:pattFill`,
	 * `a:noFill`, or `a:grpFill`. The list is ordered and 1-indexed by
	 * position; supply `orderHint` to recover document order when the
	 * parser collapsed heterogeneous siblings.
	 */
	public legacyParseFillStyleList(
		listNode: XmlObject | undefined,
		orderHint?: readonly FillListChildTag[],
	): PptxThemeFillStyle[] {
		if (!listNode) {
			return [];
		}
		return this.legacyCollectFillChildren(listNode, orderHint);
	}

	/**
	 * Parse `a:lnStyleLst` children (`a:ln` elements) into an array of
	 * {@link PptxThemeLineStyle} entries.
	 */
	public legacyParseLineStyleList(listNode: XmlObject | undefined): PptxThemeLineStyle[] {
		if (!listNode) {
			return [];
		}
		const lnNodes = this.ensureArray(listNode['a:ln']);
		return lnNodes.map((lnRaw) => {
			const ln = lnRaw as XmlObject;
			const style: PptxThemeLineStyle = { rawNode: ln };

			// Width
			if (ln['@_w']) {
				style.width = parseInt(String(ln['@_w'])) / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
			}

			// Fill colour (solid, gradient first stop, pattern foreground)
			if (ln['a:solidFill']) {
				style.color = this.parseColor(ln['a:solidFill'] as XmlObject);
				style.opacity = this.extractColorOpacity(ln['a:solidFill'] as XmlObject);
			} else if (ln['a:gradFill']) {
				style.color = this.extractGradientFillColor(ln['a:gradFill'] as XmlObject);
			} else if (ln['a:pattFill']) {
				const pf = ln['a:pattFill'] as XmlObject;
				style.color =
					this.parseColor(pf['a:fgClr'] as XmlObject | undefined) ||
					this.parseColor(pf['a:bgClr'] as XmlObject | undefined);
			}

			// Dash style
			const dashVal = (ln['a:prstDash'] as XmlObject | undefined)?.['@_val'];
			const dashType = this.normalizeStrokeDashType(dashVal);
			if (dashType) {
				style.dash = dashType;
			}

			// Line join, self-closing tags (<a:round/>) are parsed as falsy by
			// fast-xml-parser, so check key existence instead of truthiness.
			if ('a:round' in ln) {
				style.lineJoin = 'round';
			} else if ('a:bevel' in ln) {
				style.lineJoin = 'bevel';
			} else if ('a:miter' in ln) {
				style.lineJoin = 'miter';
			}

			// Line cap
			const capVal = String(ln['@_cap'] || '')
				.trim()
				.toLowerCase();
			if (capVal === 'rnd' || capVal === 'sq' || capVal === 'flat') {
				style.lineCap = capVal as PptxThemeLineStyle['lineCap'];
			}

			// Compound line
			const cmpd = String(ln['@_cmpd'] || '').trim();
			if (
				cmpd === 'sng' ||
				cmpd === 'dbl' ||
				cmpd === 'thickThin' ||
				cmpd === 'thinThick' ||
				cmpd === 'tri'
			) {
				style.compoundLine = cmpd as PptxThemeLineStyle['compoundLine'];
			}

			return style;
		});
	}

	/**
	 * Parse `a:effectStyleLst` children (`a:effectStyle`) into an array
	 * of {@link PptxThemeEffectStyle} entries.  Each style wraps an
	 * `a:effectLst` node that can contain shadow, glow, soft-edge, etc.
	 */
	public legacyParseEffectStyleList(listNode: XmlObject | undefined): PptxThemeEffectStyle[] {
		if (!listNode) {
			return [];
		}
		const styleNodes = this.ensureArray(listNode['a:effectStyle']);
		return styleNodes.map((esRaw) => {
			const es = esRaw as XmlObject;
			const effectLst = (es['a:effectLst'] ?? es['a:effectDag']) as XmlObject | undefined;
			const result: PptxThemeEffectStyle = { rawNode: es };

			if (!effectLst) {
				return result;
			}

			// Outer shadow (a:outerShdw)
			const outerShdw = effectLst['a:outerShdw'] as XmlObject | undefined;
			if (outerShdw) {
				result.shadowColor = this.parseColor(outerShdw);
				result.shadowOpacity = this.extractColorOpacity(outerShdw);
				const blurRad = parseInt(String(outerShdw['@_blurRad'] || '0'));
				if (Number.isFinite(blurRad) && blurRad > 0) {
					result.shadowBlur = blurRad / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
				const dist = parseInt(String(outerShdw['@_dist'] || '0'));
				const dir = parseInt(String(outerShdw['@_dir'] || '0'));
				if (Number.isFinite(dist) && dist > 0 && Number.isFinite(dir)) {
					const angleRad = (dir / 60000) * (Math.PI / 180);
					result.shadowOffsetX =
						(Math.cos(angleRad) * dist) / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
					result.shadowOffsetY =
						(Math.sin(angleRad) * dist) / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
			}

			// Inner shadow (a:innerShdw)
			const innerShdw = effectLst['a:innerShdw'] as XmlObject | undefined;
			if (innerShdw) {
				result.innerShadowColor = this.parseColor(innerShdw);
				result.innerShadowOpacity = this.extractColorOpacity(innerShdw);
				const blurRad = parseInt(String(innerShdw['@_blurRad'] || '0'));
				if (Number.isFinite(blurRad) && blurRad > 0) {
					result.innerShadowBlur = blurRad / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
				const dist = parseInt(String(innerShdw['@_dist'] || '0'));
				const dir = parseInt(String(innerShdw['@_dir'] || '0'));
				if (Number.isFinite(dist) && dist > 0 && Number.isFinite(dir)) {
					const angleRad = (dir / 60000) * (Math.PI / 180);
					result.innerShadowOffsetX =
						(Math.cos(angleRad) * dist) / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
					result.innerShadowOffsetY =
						(Math.sin(angleRad) * dist) / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
			}

			// Glow (a:glow)
			const glow = effectLst['a:glow'] as XmlObject | undefined;
			if (glow) {
				result.glowColor = this.parseColor(glow);
				result.glowOpacity = this.extractColorOpacity(glow);
				const glowRad = parseInt(String(glow['@_rad'] || '0'));
				if (Number.isFinite(glowRad) && glowRad > 0) {
					result.glowRadius = glowRad / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
			}

			// Soft edge (a:softEdge)
			const softEdge = effectLst['a:softEdge'] as XmlObject | undefined;
			if (softEdge) {
				const rad = parseInt(String(softEdge['@_rad'] || '0'));
				if (Number.isFinite(rad) && rad > 0) {
					result.softEdgeRadius = rad / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
			}

			// Reflection (a:reflection)
			const reflection = effectLst['a:reflection'] as XmlObject | undefined;
			if (reflection) {
				const blurRad = parseInt(String(reflection['@_blurRad'] || '0'));
				if (Number.isFinite(blurRad) && blurRad >= 0) {
					result.reflectionBlurRadius = blurRad / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
				const stA = parseInt(String(reflection['@_stA'] || ''));
				if (Number.isFinite(stA)) {
					result.reflectionStartOpacity = stA / 100000;
				}
				const endA = parseInt(String(reflection['@_endA'] || ''));
				if (Number.isFinite(endA)) {
					result.reflectionEndOpacity = endA / 100000;
				}
				const endPos = parseInt(String(reflection['@_endPos'] || ''));
				if (Number.isFinite(endPos)) {
					result.reflectionEndPosition = endPos / 100000;
				}
				const dir = parseInt(String(reflection['@_dir'] || ''));
				if (Number.isFinite(dir)) {
					result.reflectionDirection = dir / 60000;
				}
				const rot = parseInt(String(reflection['@_rot'] || ''));
				if (Number.isFinite(rot)) {
					result.reflectionRotation = rot / 60000;
				}
				const dist = parseInt(String(reflection['@_dist'] || '0'));
				if (Number.isFinite(dist) && dist >= 0) {
					result.reflectionDistance = dist / LegacyThemeFormatSchemeRuntime.EMU_PER_PX;
				}
			}

			// 3D scene (a:scene3d), sits on the effectStyle node, not effectLst
			const scene3dNode = es['a:scene3d'] as XmlObject | undefined;
			if (scene3dNode) {
				const camera = scene3dNode['a:camera'] as XmlObject | undefined;
				const lightRig = scene3dNode['a:lightRig'] as XmlObject | undefined;
				result.scene3d = {
					cameraPreset: String(camera?.['@_prst'] || '').trim() || undefined,
					lightRigType: String(lightRig?.['@_rig'] || '').trim() || undefined,
					lightRigDirection: String(lightRig?.['@_dir'] || '').trim() || undefined,
				};
			}

			// 3D shape extrusion/bevel (a:sp3d), sits on the effectStyle node
			const sp3dNode = es['a:sp3d'] as XmlObject | undefined;
			if (sp3dNode) {
				const bevelTop = sp3dNode['a:bevelT'] as XmlObject | undefined;
				const bevelBottom = sp3dNode['a:bevelB'] as XmlObject | undefined;
				result.shape3d = {
					extrusionHeight:
						sp3dNode['@_extrusionH'] !== undefined
							? parseInt(String(sp3dNode['@_extrusionH']), 10)
							: undefined,
					extrusionColor: this.parseColor(sp3dNode['a:extrusionClr'] as XmlObject | undefined),
					contourWidth:
						sp3dNode['@_contourW'] !== undefined
							? parseInt(String(sp3dNode['@_contourW']), 10)
							: undefined,
					contourColor: this.parseColor(sp3dNode['a:contourClr'] as XmlObject | undefined),
					presetMaterial: String(sp3dNode['@_prstMaterial'] || '').trim() || undefined,
					bevelTopType: bevelTop ? String(bevelTop['@_prst'] || 'circle').trim() : undefined,
					bevelTopWidth:
						bevelTop !== undefined && bevelTop['@_w'] !== undefined
							? parseInt(String(bevelTop['@_w']), 10)
							: undefined,
					bevelTopHeight:
						bevelTop !== undefined && bevelTop['@_h'] !== undefined
							? parseInt(String(bevelTop['@_h']), 10)
							: undefined,
					bevelBottomType: bevelBottom
						? String(bevelBottom['@_prst'] || 'circle').trim()
						: undefined,
					bevelBottomWidth:
						bevelBottom !== undefined && bevelBottom['@_w'] !== undefined
							? parseInt(String(bevelBottom['@_w']), 10)
							: undefined,
					bevelBottomHeight:
						bevelBottom !== undefined && bevelBottom['@_h'] !== undefined
							? parseInt(String(bevelBottom['@_h']), 10)
							: undefined,
				};
			}

			return result;
		});
	}

	/**
	 * Resolve a `a:effectRef` element into concrete effect properties
	 * by looking up `@_idx` in the theme format scheme's effect style list.
	 */
	public legacyResolveThemeEffectRef(refNode: XmlObject, style: ShapeStyle): void {
		const idx = parseInt(String(refNode['@_idx'] || '0'), 10);
		// Persist the ref data on the model so it can be re-emitted at save time
		// (Phase 2 Stream B / C-H2, preserve style refs).
		if (Number.isFinite(idx) && idx > 0) {
			style.effectRefIdx = idx;
		}
		const overrideColorXml = extractStyleReferenceColorXml(refNode);
		if (overrideColorXml) {
			style.effectRefColorXml = overrideColorXml;
		}
		if (
			!Number.isFinite(idx) ||
			idx <= 0 ||
			!this.themeFormatScheme ||
			idx > this.themeFormatScheme.effectStyles.length
		) {
			return; // No fallback for effects, they simply don't render
		}

		let effectDef = this.themeFormatScheme.effectStyles[idx - 1];
		if (!effectDef) {
			return;
		}
		const overrideColor = this.parseColor(refNode);
		if (overrideColor && effectDef.rawNode) {
			effectDef = withThemePlaceholderColor(this.themeColorMap, overrideColor, () => {
				const reparsed = this.legacyParseEffectStyleList({
					'a:effectStyle': effectDef.rawNode as XmlObject,
				});
				return reparsed[0] ?? effectDef;
			});
		}

		// Apply shadow (only if no explicit shadow was already set)
		if (effectDef.shadowColor && !style.shadowColor) {
			style.shadowColor = effectDef.shadowColor;
			style.shadowBlur = effectDef.shadowBlur;
			style.shadowOffsetX = effectDef.shadowOffsetX;
			style.shadowOffsetY = effectDef.shadowOffsetY;
			style.shadowOpacity = effectDef.shadowOpacity;
		}

		// Apply inner shadow
		if (effectDef.innerShadowColor && !style.innerShadowColor) {
			style.innerShadowColor = effectDef.innerShadowColor;
			style.innerShadowBlur = effectDef.innerShadowBlur;
			style.innerShadowOffsetX = effectDef.innerShadowOffsetX;
			style.innerShadowOffsetY = effectDef.innerShadowOffsetY;
			style.innerShadowOpacity = effectDef.innerShadowOpacity;
		}

		// Apply glow
		if (effectDef.glowColor && !style.glowColor) {
			style.glowColor = effectDef.glowColor;
			style.glowRadius = effectDef.glowRadius;
			style.glowOpacity = effectDef.glowOpacity;
		}

		// Apply soft edge
		if (effectDef.softEdgeRadius && !style.softEdgeRadius) {
			style.softEdgeRadius = effectDef.softEdgeRadius;
		}

		// Apply reflection (only if no explicit reflection was already set)
		if (effectDef.reflectionBlurRadius !== undefined && style.reflectionBlurRadius === undefined) {
			style.reflectionBlurRadius = effectDef.reflectionBlurRadius;
			style.reflectionStartOpacity = effectDef.reflectionStartOpacity;
			style.reflectionEndOpacity = effectDef.reflectionEndOpacity;
			style.reflectionEndPosition = effectDef.reflectionEndPosition;
			style.reflectionDirection = effectDef.reflectionDirection;
			style.reflectionRotation = effectDef.reflectionRotation;
			style.reflectionDistance = effectDef.reflectionDistance;
		}

		// Apply 3D scene (only if no explicit scene3d was already set)
		if (effectDef.scene3d && !style.scene3d) {
			style.scene3d = effectDef.scene3d;
		}

		// Apply 3D shape extrusion/bevel (only if no explicit shape3d was already set)
		if (effectDef.shape3d && !style.shape3d) {
			style.shape3d = effectDef.shape3d;
		}
	}

	/**
	 * Resolve a `a:lnRef` element into concrete stroke properties by
	 * looking up `@_idx` in the theme format scheme's line style list.
	 */
	public legacyResolveThemeLineRef(refNode: XmlObject, style: ShapeStyle): void {
		const idx = parseInt(String(refNode['@_idx'] || '0'), 10);
		const overrideColor = this.parseColor(refNode);
		// Persist the ref data on the model so it can be re-emitted at save time.
		if (Number.isFinite(idx) && idx > 0) {
			style.lnRefIdx = idx;
		}
		const overrideColorXml = extractStyleReferenceColorXml(refNode);
		if (overrideColorXml) {
			style.lnRefColorXml = overrideColorXml;
		}
		// `idx="0"` references no line style at all: PowerPoint draws no
		// outline (COM: Shape Styles "Intense Effect" row, Line.Visible = 0),
		// exactly as `fillRef idx="0"` means no fill. The colour child only
		// matters for a real list entry.
		if (idx === 0 && String(refNode['@_idx'] ?? '').trim() === '0') {
			style.strokeWidth = 0;
			style.strokeColor = 'transparent';
			return;
		}

		if (
			!Number.isFinite(idx) ||
			idx <= 0 ||
			!this.themeFormatScheme ||
			idx > this.themeFormatScheme.lineStyles.length
		) {
			// Fallback to pre-existing behaviour
			style.strokeColor = overrideColor;
			if (overrideColor) {
				style.strokeWidth = 1;
			}
			return;
		}

		let lineDef = this.themeFormatScheme.lineStyles[idx - 1];
		if (!lineDef) {
			style.strokeColor = overrideColor;
			if (overrideColor) {
				style.strokeWidth = 1;
			}
			return;
		}
		if (overrideColor && lineDef.rawNode) {
			lineDef = withThemePlaceholderColor(this.themeColorMap, overrideColor, () => {
				const reparsed = this.legacyParseLineStyleList({ 'a:ln': lineDef.rawNode as XmlObject });
				return reparsed[0] ?? lineDef;
			});
		}

		// Apply line properties from the format scheme
		style.strokeColor = lineDef.color || overrideColor;
		if (lineDef.opacity !== undefined) {
			style.strokeOpacity = lineDef.opacity;
		}
		if (lineDef.width !== undefined && lineDef.width > 0) {
			style.strokeWidth = lineDef.width;
		} else if (style.strokeColor) {
			style.strokeWidth = 1;
		}
		if (lineDef.dash) {
			style.strokeDash = lineDef.dash as ShapeStyle['strokeDash'];
		}
		if (lineDef.lineJoin) {
			style.lineJoin = lineDef.lineJoin;
		}
		if (lineDef.lineCap) {
			style.lineCap = lineDef.lineCap;
		}
		if (lineDef.compoundLine) {
			style.compoundLine = lineDef.compoundLine;
		}
	}

	/**
	 * Re-resolve a gradient fill's stops by substituting `phClr` occurrences
	 * with the given override colour.  When the gradient node contains
	 * `a:schemeClr @val="phClr"`, the override colour is injected.
	 */
	public legacyReResolveGradientWithPhClr(
		gradNode: XmlObject,
		phClrValue: string,
	): {
		stops: NonNullable<ShapeStyle['fillGradientStops']>;
		css: string | undefined;
	} {
		return withThemePlaceholderColor(this.themeColorMap, phClrValue, () => {
			const stops = this.extractGradientStops(gradNode);
			const type = this.extractGradientType(gradNode);
			const angle = this.extractGradientAngle(gradNode);
			return { stops, css: this.buildGradientCssFromStops(stops, type, angle) };
		});
	}

	/**
	 * Resolve a `a:fillRef` element into concrete {@link ShapeStyle} fill
	 * properties by looking up the 1-based `@_idx` in the theme's format
	 * scheme fill style list.
	 *
	 * Per the OOXML spec:
	 * - idx 0 or 1000 → no fill
	 * - idx 1-3 → fillStyleLst[idx-1]
	 * - idx 1001-1003 → bgFillStyleLst[idx-1001]
	 *
	 * The colour child of the ref (e.g. `a:schemeClr`) acts as a
	 * "placeholder colour" (`phClr`) override for `phClr` tokens inside
	 * the fill definition.
	 */
	public legacyResolveThemeFillRef(refNode: XmlObject, style: ShapeStyle): void {
		const idx = parseInt(String(refNode['@_idx'] || '0'), 10);
		// Persist the ref data on the model so it can be re-emitted at save time.
		if (Number.isFinite(idx) && idx > 0) {
			style.fillRefIdx = idx;
		}
		const overrideColorXml = extractStyleReferenceColorXml(refNode);
		if (overrideColorXml) {
			style.fillRefColorXml = overrideColorXml;
		}
		if (idx === 0 || idx === 1000) {
			style.fillMode = 'none';
			style.fillColor = 'transparent';
			style.fillOpacity = 0;
			return;
		}
		if (!Number.isFinite(idx) || idx <= 0 || !this.themeFormatScheme) {
			// Fallback: just use the colour child (pre-existing behaviour)
			style.fillMode = 'theme';
			style.fillColor = this.parseColor(refNode);
			return;
		}

		// Determine which list and which 0-based offset
		let fillDef: PptxThemeFillStyle | undefined;
		if (idx >= 1001) {
			const offset = idx - 1001;
			fillDef = this.themeFormatScheme.backgroundFillStyles[offset];
		} else {
			fillDef = this.themeFormatScheme.fillStyles[idx - 1];
		}

		if (!fillDef) {
			style.fillMode = 'theme';
			style.fillColor = this.parseColor(refNode);
			return;
		}

		// The colour child of the ref element serves as phClr override.
		// We resolve it for use when the fill definition contains phClr tokens.
		const overrideColor = this.parseColor(refNode);
		const fillTag =
			fillDef.kind === 'solid'
				? 'a:solidFill'
				: fillDef.kind === 'pattern'
					? 'a:pattFill'
					: undefined;
		if (fillTag && overrideColor && fillDef.rawNode) {
			const rawFillNode = fillDef.rawNode as XmlObject;
			fillDef = withThemePlaceholderColor(this.themeColorMap, overrideColor, () => {
				const reparsed = this.legacyParseFillStyleList({ [fillTag]: rawFillNode } as XmlObject);
				return reparsed[0] ?? fillDef;
			});
		}

		switch (fillDef.kind) {
			case 'solid': {
				style.fillMode = 'solid';
				style.fillColor = fillDef.color || overrideColor;
				style.fillOpacity = fillDef.opacity;
				break;
			}
			case 'gradient': {
				style.fillMode = 'gradient';
				// For phClr gradients re-resolve stops with the override colour.
				if (fillDef.rawNode && overrideColor) {
					const reresolved = this.legacyReResolveGradientWithPhClr(
						fillDef.rawNode as XmlObject,
						overrideColor,
					);
					style.fillGradientStops = reresolved.stops;
					style.fillGradient = reresolved.css;
				} else {
					style.fillGradientStops = fillDef.gradientStops;
					style.fillGradient = fillDef.gradientCss;
				}
				style.fillGradientAngle = fillDef.gradientAngle;
				style.fillGradientType = fillDef.gradientType;
				style.fillColor = fillDef.color || overrideColor;
				break;
			}
			case 'pattern': {
				style.fillMode = 'pattern';
				style.fillColor = fillDef.color || overrideColor;
				if (fillDef.patternPreset) {
					style.fillPatternPreset = fillDef.patternPreset;
				}
				if (fillDef.patternBackgroundColor) {
					style.fillPatternBackgroundColor = fillDef.patternBackgroundColor;
				}
				break;
			}
			case 'none': {
				style.fillMode = 'none';
				style.fillColor = 'transparent';
				style.fillOpacity = 0;
				break;
			}
		}
	}
	/**
	 * Parse the `a:fmtScheme` element from the theme into a structured
	 * {@link PptxThemeFormatScheme}.  Each sub-list (fillStyleLst, lnStyleLst,
	 * effectStyleLst, bgFillStyleLst) contains up to three intensity levels
	 * (subtle / moderate / intense) referenced by 1-based index.
	 *
	 * When `rawXml` is supplied (the original theme/themeOverride XML
	 * source), the document order of fill children is recovered via a
	 * targeted regex scan and threaded into {@link legacyParseFillStyleList} so
	 * mixed solid/grad/blip/grpFill lists round-trip in source order even
	 * though fast-xml-parser collapses heterogeneous siblings into typed
	 * buckets.
	 */
	public legacyParseFormatScheme(fmtScheme: XmlObject, rawXml?: string): PptxThemeFormatScheme {
		const name = String(fmtScheme['@_name'] || '').trim() || undefined;

		const fillOrder = extractFillStyleListChildOrder(rawXml, 'a:fillStyleLst');
		const bgFillOrder = extractFillStyleListChildOrder(rawXml, 'a:bgFillStyleLst');

		// --- Fill styles (a:fillStyleLst) ---
		const fillStyleLst = fmtScheme['a:fillStyleLst'] as XmlObject | undefined;
		const fillStyles = this.legacyParseFillStyleList(fillStyleLst, fillOrder);

		// --- Line styles (a:lnStyleLst) ---
		const lnStyleLst = fmtScheme['a:lnStyleLst'] as XmlObject | undefined;
		const lineStyles = this.legacyParseLineStyleList(lnStyleLst);

		// --- Effect styles (a:effectStyleLst) ---
		const effectStyleLst = fmtScheme['a:effectStyleLst'] as XmlObject | undefined;
		const effectStyles = this.legacyParseEffectStyleList(effectStyleLst);

		// --- Background fill styles (a:bgFillStyleLst) ---
		const bgFillStyleLst = fmtScheme['a:bgFillStyleLst'] as XmlObject | undefined;
		const backgroundFillStyles = this.legacyParseFillStyleList(bgFillStyleLst, bgFillOrder);

		return {
			name,
			fillStyles,
			lineStyles,
			effectStyles,
			backgroundFillStyles,
		};
	}

	// The theme part writer (`PptxHandlerRuntimeSaveTheme`) before it moved onto `drawingThemeXml`.
	/**
	 * Build a complete `<a:theme>` XML document from in-memory state.
	 * Returns the serialized XML string (with XML prolog), or `undefined`
	 * if there is no source data to emit.
	 *
	 * - Color scheme: built from per-master color map (or global fallback).
	 * - Font scheme: built from per-master font map + per-script entries.
	 * - Format scheme: re-emit the original XML subtree if available; else
	 *   build a minimal scheme from {@link themeFormatScheme}.
	 * - objectDefaults / extraClrSchemeLst / custClrLst / extLst: re-emit
	 *   captured raw subtrees.
	 */
	public legacyBuildThemeXml(
		themePath: string,
		masterPath: string | undefined,
	): string | undefined {
		const colorMap =
			(masterPath && this.masterThemeColorMaps.get(masterPath)) ||
			this.globalThemeColorMapSnapshot ||
			this.themeColorMap;
		const fontMap =
			(masterPath && this.masterThemeFontMaps.get(masterPath)) ||
			this.globalThemeFontMapSnapshot ||
			this.themeFontMap;

		const themeName = this.masterThemeNames.get(themePath) || 'Office Theme';
		const colorSchemeName = this.masterThemeColorSchemeNames.get(themePath) || themeName;
		const fontSchemeName = this.masterThemeFontSchemeNames.get(themePath) || themeName;

		const majorScripts = this.masterThemeMajorFontScripts.get(themePath) || {};
		const minorScripts = this.masterThemeMinorFontScripts.get(themePath) || {};

		// Build colour scheme XML object.
		const clrScheme = this.legacyBuildClrSchemeObject(colorSchemeName, colorMap);
		const fontScheme = this.legacyBuildFontSchemeObject(
			fontSchemeName,
			fontMap,
			majorScripts,
			minorScripts,
		);

		// Format scheme: prefer the original raw subtree (kept on the parsed
		// theme via originalThemeXmlByPath) since the in-memory
		// PptxThemeFormatScheme is a lossy representation. Re-extract it from
		// the original XML by parsing once and grabbing `a:fmtScheme` verbatim.
		const fmtScheme = this.legacyExtractRawSubtreeFromOriginal(themePath, [
			'a:theme',
			'a:themeElements',
			'a:fmtScheme',
		]);

		const themeElements: XmlObject = {
			'a:clrScheme': clrScheme,
			'a:fontScheme': fontScheme,
		};
		if (fmtScheme !== undefined) {
			themeElements['a:fmtScheme'] = fmtScheme as XmlObject;
		} else {
			// Last-resort minimal placeholder so PowerPoint doesn't reject
			// the theme. (Required child per CT_BaseStyles.)
			themeElements['a:fmtScheme'] = this.legacyBuildMinimalFmtScheme(themeName);
		}

		const themeRoot: XmlObject = {
			'@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
			'@_name': themeName,
			'a:themeElements': themeElements,
		};

		// objectDefaults, M5.
		const objectDefaults = this.masterThemeObjectDefaults.get(themePath);
		if (objectDefaults && (objectDefaults.spDef || objectDefaults.lnDef || objectDefaults.txDef)) {
			const od: XmlObject = {};
			if (objectDefaults.spDef !== undefined) {
				od['a:spDef'] = objectDefaults.spDef as XmlObject;
			}
			if (objectDefaults.lnDef !== undefined) {
				od['a:lnDef'] = objectDefaults.lnDef as XmlObject;
			}
			if (objectDefaults.txDef !== undefined) {
				od['a:txDef'] = objectDefaults.txDef as XmlObject;
			}
			themeRoot['a:objectDefaults'] = od;
		} else {
			// Empty objectDefaults is canonical per Office output.
			themeRoot['a:objectDefaults'] = {};
		}

		// extraClrSchemeLst.
		const extraClr = this.masterThemeExtraClrSchemeLst.get(themePath);
		themeRoot['a:extraClrSchemeLst'] =
			extraClr !== undefined ? (extraClr as XmlObject) : ({} as XmlObject);

		// custClrLst.
		const custClr = this.masterThemeCustClrLst.get(themePath);
		if (custClr !== undefined) {
			themeRoot['a:custClrLst'] = custClr as XmlObject;
		}

		// extLst.
		const themeExt = this.masterThemeExtLst.get(themePath);
		if (themeExt !== undefined) {
			themeRoot['a:extLst'] = themeExt as XmlObject;
		}

		const doc: XmlObject = {
			'?xml': { '@_version': '1.0', '@_encoding': 'UTF-8', '@_standalone': 'yes' },
			'a:theme': themeRoot,
		};

		try {
			return this.builder.build(doc);
		} catch (error) {
			console.warn(`Failed to build theme XML for ${themePath}:`, error);
			return undefined;
		}
	}

	/**
	 * Build the `a:clrScheme` XmlObject from a colour map. Each slot
	 * value is interpreted as either a `#RRGGBB` srgb hex or a known
	 * sysClr token (currently always emitted as srgbClr, the in-memory
	 * map is hex-typed; sysClr round-trip belongs to the broader C-H3
	 * fix to preserve original color XML and is out of scope here).
	 */
	public legacyBuildClrSchemeObject(
		schemeName: string,
		colorMap: Record<string, string>,
	): XmlObject {
		const slot = (key: keyof PptxThemeColorScheme): XmlObject => {
			const hex = String(colorMap[key as string] || '').replace(/^#/, '');
			const srgb = hex.length === 6 ? hex.toUpperCase() : '000000';
			return { 'a:srgbClr': { '@_val': srgb } };
		};

		// dk1/lt1 typically use sysClr in real Office themes (windowText /
		// window). When the original theme's <a:dk1> (or <a:lt1>) was a
		// sysClr we lose that nuance through the hex-only colour map. Emit
		// srgbClr unconditionally; downstream consumers that load this file
		// see the same resolved colour values.
		return {
			'@_name': schemeName,
			'a:dk1': slot('dk1'),
			'a:lt1': slot('lt1'),
			'a:dk2': slot('dk2'),
			'a:lt2': slot('lt2'),
			'a:accent1': slot('accent1'),
			'a:accent2': slot('accent2'),
			'a:accent3': slot('accent3'),
			'a:accent4': slot('accent4'),
			'a:accent5': slot('accent5'),
			'a:accent6': slot('accent6'),
			'a:hlink': slot('hlink'),
			'a:folHlink': slot('folHlink'),
		};
	}

	/**
	 * Build the `a:fontScheme` XmlObject from a font map plus per-script
	 * font tables.
	 *
	 * Phase 4 Stream A / M4.
	 */
	public legacyBuildFontSchemeObject(
		schemeName: string,
		fontMap: Record<string, string>,
		majorScripts: Record<string, string>,
		minorScripts: Record<string, string>,
	): XmlObject {
		const buildFontGroup = (
			latinKey: string,
			eaKey: string,
			csKey: string,
			scripts: Record<string, string>,
		): XmlObject => {
			const group: XmlObject = {
				'a:latin': { '@_typeface': fontMap[latinKey] || 'Calibri' },
				'a:ea': { '@_typeface': fontMap[eaKey] || '' },
				'a:cs': { '@_typeface': fontMap[csKey] || '' },
			};
			const scriptKeys = Object.keys(scripts);
			if (scriptKeys.length > 0) {
				const fontEntries = scriptKeys.map((script) => ({
					'@_script': script,
					'@_typeface': scripts[script],
				}));
				group['a:font'] = fontEntries.length === 1 ? fontEntries[0] : fontEntries;
			}
			return group;
		};

		return {
			'@_name': schemeName,
			'a:majorFont': buildFontGroup('mj-lt', 'mj-ea', 'mj-cs', majorScripts),
			'a:minorFont': buildFontGroup('mn-lt', 'mn-ea', 'mn-cs', minorScripts),
		};
	}

	/**
	 * Re-parse the original theme XML and pluck out a subtree by path,
	 * returning the raw parser object. Returns `undefined` when the
	 * original is missing or the path doesn't exist.
	 *
	 * Used to preserve `a:fmtScheme` byte-for-byte through a regenerate
	 * round-trip, since the in-memory PptxThemeFormatScheme is lossy.
	 */
	public legacyExtractRawSubtreeFromOriginal(
		themePath: string,
		path: string[],
	): unknown | undefined {
		const original = this.originalThemeXmlByPath.get(themePath);
		if (!original) {
			return undefined;
		}
		try {
			const parsed = this.parser.parse(original) as XmlObject;
			let cursor: unknown = parsed;
			for (const segment of path) {
				if (cursor && typeof cursor === 'object' && segment in (cursor as XmlObject)) {
					cursor = (cursor as XmlObject)[segment];
				} else {
					return undefined;
				}
			}
			return cursor;
		} catch {
			return undefined;
		}
	}

	/**
	 * Last-resort minimal `<a:fmtScheme>` body. Mirrors the SDK new-deck
	 * builder's output for new presentations, scaled down to the smallest
	 * schema-valid form.
	 */
	public legacyBuildMinimalFmtScheme(name: string): XmlObject {
		const phClrSolid = { 'a:solidFill': { 'a:schemeClr': { '@_val': 'phClr' } } };
		return {
			'@_name': name,
			'a:fillStyleLst': {
				'a:solidFill': [{ 'a:schemeClr': { '@_val': 'phClr' } }],
				'a:gradFill': [],
			},
			'a:lnStyleLst': {
				'a:ln': [
					{
						'@_w': '6350',
						'@_cap': 'flat',
						'@_cmpd': 'sng',
						'@_algn': 'ctr',
						...phClrSolid,
						'a:prstDash': { '@_val': 'solid' },
						'a:miter': { '@_lim': '800000' },
					},
					{
						'@_w': '12700',
						'@_cap': 'flat',
						'@_cmpd': 'sng',
						'@_algn': 'ctr',
						...phClrSolid,
						'a:prstDash': { '@_val': 'solid' },
						'a:miter': { '@_lim': '800000' },
					},
					{
						'@_w': '19050',
						'@_cap': 'flat',
						'@_cmpd': 'sng',
						'@_algn': 'ctr',
						...phClrSolid,
						'a:prstDash': { '@_val': 'solid' },
						'a:miter': { '@_lim': '800000' },
					},
				],
			},
			'a:effectStyleLst': {
				'a:effectStyle': [{ 'a:effectLst': {} }, { 'a:effectLst': {} }, { 'a:effectLst': {} }],
			},
			'a:bgFillStyleLst': {
				'a:solidFill': [{ 'a:schemeClr': { '@_val': 'phClr' } }],
				'a:gradFill': [],
			},
		};
	}
}
