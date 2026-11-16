/**
 * @fileoverview The pptx view of a theme format scheme (`a:fmtScheme`), projected from the shared
 * DrawingML reader (`parseThemeFormatScheme`, ECMA-376 Part 1, 20.1.4.1.14).
 *
 * The neutral {@link ThemeFormatScheme} is the source of truth: list order (a fill list in
 * document order, an `a:blipFill` kept at its index), fill kinds, colours, pattern presets, line
 * geometry and effect geometry come from it, and the style references resolve against it with the
 * shared `resolveTheme*Style` helpers, which also do the `phClr` substitution. What the pptx
 * object tree still supplies is what other codecs own: the gradient view (stops, CSS, angle) of
 * the gradient codec, the colour opacity reader and the `a:scene3d`/`a:sp3d` 3D settings, read
 * from the object node paired with each entry (`rawNode`, which table styles and backgrounds use).
 * See docs/agnostic-core-plan.md, step 3.
 *
 * @module theme-format-scheme-view
 */
import type { DrawingEffect, DrawingEffectList } from '../../../../drawingml/effect-model';
import type {
	ThemeEffectStyle,
	ThemeFillStyle,
	ThemeFormatScheme,
} from '../../../../drawingml/format-scheme';
import type { DrawingColor, DrawingFill, DrawingLine } from '../../../../drawingml/types';
import { NS, elements, parseXml, type XmlElement } from '../../../../xml/index';
import type {
	PptxThemeEffectStyle,
	PptxThemeFillStyle,
	PptxThemeFormatScheme,
	PptxThemeLineStyle,
	XmlObject,
} from '../../types';

/** What the projection needs of the runtime. */
export interface ThemeStyleViewContext {
	/** `#RRGGBB` of a neutral colour through the active theme (`phClr` already substituted). */
	color(color: DrawingColor | undefined): string | undefined;
	/** The opacity the object-tree colour holder declares (`a:alpha`...), if any. */
	opacity(holder: XmlObject | undefined): number | undefined;
	/** The gradient view of an object-tree `a:gradFill`. */
	gradient(
		node: XmlObject,
	): Pick<
		PptxThemeFillStyle,
		'color' | 'opacity' | 'gradientStops' | 'gradientAngle' | 'gradientType' | 'gradientCss'
	>;
	/** The representative colour of an object-tree `a:gradFill`. */
	gradientColor(node: XmlObject): string | undefined;
	/** `#RRGGBB` of an object-tree colour holder (the 3D colours). */
	objectColor(holder: XmlObject | undefined): string | undefined;
	normalizeDash(value: string | undefined): string | undefined;
	emuPerPx: number;
}

const asObject = (value: unknown): XmlObject | undefined =>
	value && typeof value === 'object' ? (value as XmlObject) : undefined;

/** The `index`th object-tree child named `name` (same-name siblings arrive as an array). */
export function objectChildAt(node: unknown, name: string, index: number): XmlObject | undefined {
	const holder = asObject(node);
	if (!holder || !(name in holder)) return undefined;
	const value = holder[name];
	if (Array.isArray(value)) return value[index] === undefined ? undefined : toNode(value[index]);
	return index === 0 ? toNode(value) : undefined;
}

/** A self-closing element arrives as `''`; it is an empty node. */
const toNode = (value: unknown): XmlObject => asObject(value) ?? {};

const PREFIX_PATTERN = /<\/?([A-Za-z_][\w.-]*):|\s([A-Za-z_][\w.-]*):[\w.-]+\s*=/g;

/**
 * The DOM element of an object-tree node, for the shared readers. `build` is the runtime's XML
 * builder; every prefix the text uses is declared (DrawingML's `a` with its real URI, the rest
 * with stand-ins the DrawingML readers ignore). `undefined` when the text does not parse.
 */
export function objectNodeElement(
	build: (tree: XmlObject) => string,
	tag: string,
	node: XmlObject,
): XmlElement | undefined {
	const xml = build({ [tag]: node } as XmlObject);
	const prefixes = new Set<string>(['a']);
	for (const match of xml.matchAll(PREFIX_PATTERN)) {
		const prefix = match[1] ?? match[2];
		if (prefix && prefix !== 'xmlns' && prefix !== 'xml') prefixes.add(prefix);
	}
	const declarations = [...prefixes]
		.map((prefix) => {
			const uri = (NS as Record<string, string>)[prefix] ?? `urn:ooxml-core:prefix:${prefix}`;
			return `xmlns:${prefix}="${uri}"`;
		})
		.join(' ');
	try {
		const root = parseXml(`<w ${declarations}>${xml}</w>`).documentElement;
		return elements(root)[0];
	} catch {
		return undefined;
	}
}

function fillView(
	fill: DrawingFill,
	node: XmlObject | undefined,
	context: ThemeStyleViewContext,
): PptxThemeFillStyle {
	const rawNode = node ?? {};
	switch (fill.kind) {
		case 'solid':
			return {
				kind: 'solid',
				color: context.color(fill.color),
				opacity: context.opacity(node),
				fill,
				rawNode,
			};
		case 'gradient':
			return { kind: 'gradient', ...context.gradient(rawNode), fill, rawNode };
		case 'pattern':
			return {
				kind: 'pattern',
				color: context.color(fill.foreground) || context.color(fill.background),
				patternPreset: fill.preset.trim() || undefined,
				patternBackgroundColor: context.color(fill.background) || undefined,
				fill,
				rawNode,
			};
		case 'none':
			return { kind: 'none', fill, rawNode };
		case 'picture':
			return { kind: 'picture', fill, rawNode };
		case 'unsupported':
			// `a:grpFill` (the group's fill) and a colourless `a:solidFill`.
			return fill.element === 'grpFill'
				? { kind: 'group', fill, rawNode }
				: {
						kind: 'solid',
						color: undefined,
						opacity: context.opacity(node),
						fill,
						rawNode,
					};
	}
}

const FILL_TAG: Record<string, string> = {
	solid: 'a:solidFill',
	gradient: 'a:gradFill',
	pattern: 'a:pattFill',
	none: 'a:noFill',
	picture: 'a:blipFill',
};

function fillTagOf(fill: DrawingFill): string {
	return fill.kind === 'unsupported' ? `a:${fill.element}` : FILL_TAG[fill.kind];
}

/** One fill style entry (`a:fillStyleLst` or `a:bgFillStyleLst` child) as the pptx view. */
export function themeFillStyleView(
	style: ThemeFillStyle,
	node: XmlObject | undefined,
	context: ThemeStyleViewContext,
): PptxThemeFillStyle {
	return fillView(style.fill, node, context);
}

const LINE_CAPS = new Set(['rnd', 'sq', 'flat']);
const COMPOUND_LINES = new Set(['sng', 'dbl', 'thickThin', 'thinThick', 'tri']);

/** One `a:ln` of `a:lnStyleLst` as the pptx view; `node` is its object-tree `a:ln`. */
export function themeLineStyleView(
	line: DrawingLine,
	node: XmlObject | undefined,
	context: ThemeStyleViewContext,
): PptxThemeLineStyle {
	const ln = node ?? {};
	const style: PptxThemeLineStyle = { line, rawNode: ln };
	if (line.widthEmu !== undefined && ln['@_w']) style.width = line.widthEmu / context.emuPerPx;
	const fill = line.fill;
	if (fill?.kind === 'solid') {
		style.color = context.color(fill.color);
		style.opacity = context.opacity(asObject(ln['a:solidFill']));
	} else if (fill?.kind === 'gradient') {
		style.color = context.gradientColor(toNode(ln['a:gradFill']));
	} else if (fill?.kind === 'pattern') {
		style.color = context.color(fill.foreground) || context.color(fill.background);
	}
	const dash = context.normalizeDash(line.dash);
	if (dash) style.dash = dash;
	if (line.join) style.lineJoin = line.join.kind;
	const cap = line.cap?.trim().toLowerCase();
	if (cap && LINE_CAPS.has(cap)) style.lineCap = cap as PptxThemeLineStyle['lineCap'];
	const compound = line.compound?.trim();
	if (compound && COMPOUND_LINES.has(compound)) {
		style.compoundLine = compound as PptxThemeLineStyle['compoundLine'];
	}
	return style;
}

type EffectOf<K extends DrawingEffect['kind']> = Extract<DrawingEffect, { kind: K }>;

function firstEffect<K extends DrawingEffect['kind']>(
	list: DrawingEffectList | undefined,
	kind: K,
): EffectOf<K> | undefined {
	return list?.effects.find((effect): effect is EffectOf<K> => effect.kind === kind);
}

function shadowOffset(
	distanceEmu: number | undefined,
	directionAngle: number | undefined,
	emuPerPx: number,
): { x: number; y: number } | undefined {
	const dist = Math.trunc(distanceEmu ?? 0);
	const dir = Math.trunc(directionAngle ?? 0);
	if (!(dist > 0)) return undefined;
	const angle = (dir / 60000) * (Math.PI / 180);
	return { x: (Math.cos(angle) * dist) / emuPerPx, y: (Math.sin(angle) * dist) / emuPerPx };
}

const positive = (value: number | undefined): number | undefined => {
	const whole = Math.trunc(value ?? 0);
	return whole > 0 ? whole : undefined;
};

const integer = (value: number | undefined): number | undefined =>
	value === undefined ? undefined : Math.trunc(value);

function attributeInteger(node: XmlObject | undefined, name: string): number | undefined {
	const value = node?.[`@_${name}`];
	return value !== undefined ? parseInt(String(value), 10) : undefined;
}

/** The 3D settings an `a:effectStyle` carries next to its effect list (object tree). */
function effectStyle3d(
	node: XmlObject | undefined,
	context: ThemeStyleViewContext,
): Pick<PptxThemeEffectStyle, 'scene3d' | 'shape3d'> {
	const result: Pick<PptxThemeEffectStyle, 'scene3d' | 'shape3d'> = {};
	const scene3d = asObject(node?.['a:scene3d']);
	if (scene3d) {
		const camera = asObject(scene3d['a:camera']);
		const lightRig = asObject(scene3d['a:lightRig']);
		result.scene3d = {
			cameraPreset: String(camera?.['@_prst'] || '').trim() || undefined,
			lightRigType: String(lightRig?.['@_rig'] || '').trim() || undefined,
			lightRigDirection: String(lightRig?.['@_dir'] || '').trim() || undefined,
		};
	}
	const sp3d = asObject(node?.['a:sp3d']);
	if (sp3d) {
		const bevelTop = sp3d['a:bevelT'] as XmlObject | undefined;
		const bevelBottom = sp3d['a:bevelB'] as XmlObject | undefined;
		const bevelType = (bevel: XmlObject | undefined) =>
			bevel !== undefined ? String(asObject(bevel)?.['@_prst'] || 'circle').trim() : undefined;
		result.shape3d = {
			extrusionHeight: attributeInteger(sp3d, 'extrusionH'),
			extrusionColor: context.objectColor(asObject(sp3d['a:extrusionClr'])),
			contourWidth: attributeInteger(sp3d, 'contourW'),
			contourColor: context.objectColor(asObject(sp3d['a:contourClr'])),
			presetMaterial: String(sp3d['@_prstMaterial'] || '').trim() || undefined,
			bevelTopType: bevelTop ? bevelType(bevelTop) : undefined,
			bevelTopWidth: attributeInteger(asObject(bevelTop), 'w'),
			bevelTopHeight: attributeInteger(asObject(bevelTop), 'h'),
			bevelBottomType: bevelBottom ? bevelType(bevelBottom) : undefined,
			bevelBottomWidth: attributeInteger(asObject(bevelBottom), 'w'),
			bevelBottomHeight: attributeInteger(asObject(bevelBottom), 'h'),
		};
	}
	return result;
}

/** One `a:effectStyle` as the pptx view; `node` is its object-tree `a:effectStyle`. */
export function themeEffectStyleView(
	style: ThemeEffectStyle,
	node: XmlObject | undefined,
	context: ThemeStyleViewContext,
): PptxThemeEffectStyle {
	const es = node ?? {};
	const result: PptxThemeEffectStyle = {
		...(style.effects ? { effects: style.effects } : {}),
		rawNode: es,
	};
	const list = style.effects;
	const lst = asObject(es['a:effectLst']);
	if (!list) return { ...result, ...effectStyle3d(es, context) };
	const emuPerPx = context.emuPerPx;

	const outer = firstEffect(list, 'outerShadow');
	if (outer && lst?.['a:outerShdw']) {
		result.shadowColor = context.color(outer.color);
		result.shadowOpacity = context.opacity(asObject(lst['a:outerShdw']));
		const blur = positive(outer.blurRadiusEmu);
		if (blur) result.shadowBlur = blur / emuPerPx;
		const offset = shadowOffset(outer.distanceEmu, outer.directionAngle, emuPerPx);
		if (offset) {
			result.shadowOffsetX = offset.x;
			result.shadowOffsetY = offset.y;
		}
	}
	const inner = firstEffect(list, 'innerShadow');
	if (inner && lst?.['a:innerShdw']) {
		result.innerShadowColor = context.color(inner.color);
		result.innerShadowOpacity = context.opacity(asObject(lst['a:innerShdw']));
		const blur = positive(inner.blurRadiusEmu);
		if (blur) result.innerShadowBlur = blur / emuPerPx;
		const offset = shadowOffset(inner.distanceEmu, inner.directionAngle, emuPerPx);
		if (offset) {
			result.innerShadowOffsetX = offset.x;
			result.innerShadowOffsetY = offset.y;
		}
	}
	const glow = firstEffect(list, 'glow');
	if (glow && lst?.['a:glow']) {
		result.glowColor = context.color(glow.color);
		result.glowOpacity = context.opacity(asObject(lst['a:glow']));
		const radius = positive(glow.radiusEmu);
		if (radius) result.glowRadius = radius / emuPerPx;
	}
	const softEdge = firstEffect(list, 'softEdge');
	if (softEdge && lst?.['a:softEdge']) {
		const radius = positive(softEdge.radiusEmu);
		if (radius) result.softEdgeRadius = radius / emuPerPx;
	}
	const reflection = firstEffect(list, 'reflection');
	if (reflection && lst?.['a:reflection']) {
		// An absent radius or distance is zero; a negative one is not read.
		const blur = integer(reflection.blurRadiusEmu) ?? 0;
		if (blur >= 0) result.reflectionBlurRadius = blur / emuPerPx;
		const fraction = (value: number | undefined, scale: number) => {
			const whole = integer(value);
			return whole === undefined ? undefined : whole / scale;
		};
		const startOpacity = fraction(reflection.startAlpha, 100000);
		if (startOpacity !== undefined) result.reflectionStartOpacity = startOpacity;
		const endOpacity = fraction(reflection.endAlpha, 100000);
		if (endOpacity !== undefined) result.reflectionEndOpacity = endOpacity;
		const endPosition = fraction(reflection.endPosition, 100000);
		if (endPosition !== undefined) result.reflectionEndPosition = endPosition;
		const direction = fraction(reflection.directionAngle, 60000);
		if (direction !== undefined) result.reflectionDirection = direction;
		const distance = integer(reflection.distanceEmu) ?? 0;
		if (distance >= 0) result.reflectionDistance = distance / emuPerPx;
	}
	return { ...result, ...effectStyle3d(es, context) };
}

/** The object-tree nodes of a style list's entries, paired with the DOM order of `entries`. */
function pairedNodes(listNode: unknown, tags: readonly string[]): (XmlObject | undefined)[] {
	const cursors = new Map<string, number>();
	return tags.map((tag) => {
		const index = cursors.get(tag) ?? 0;
		cursors.set(tag, index + 1);
		return objectChildAt(listNode, tag, index);
	});
}

/** The pptx view of a neutral format scheme; `fmtScheme` is the same element's object tree. */
export function themeFormatSchemeView(
	scheme: ThemeFormatScheme,
	fmtScheme: XmlObject | undefined,
	context: ThemeStyleViewContext,
): PptxThemeFormatScheme {
	const fills = (styles: ThemeFillStyle[], list: string) => {
		const nodes = pairedNodes(
			fmtScheme?.[list],
			styles.map((style) => fillTagOf(style.fill)),
		);
		return styles.map((style, index) => themeFillStyleView(style, nodes[index], context));
	};
	const lineNodes = pairedNodes(
		fmtScheme?.['a:lnStyleLst'],
		scheme.lineStyles.map(() => 'a:ln'),
	);
	const effectNodes = pairedNodes(
		fmtScheme?.['a:effectStyleLst'],
		scheme.effectStyles.map(() => 'a:effectStyle'),
	);
	return {
		name: scheme.name?.trim() || undefined,
		fillStyles: fills(scheme.fillStyles, 'a:fillStyleLst'),
		lineStyles: scheme.lineStyles.map((line, index) =>
			themeLineStyleView(line, lineNodes[index], context),
		),
		effectStyles: scheme.effectStyles.map((style, index) =>
			themeEffectStyleView(style, effectNodes[index], context),
		),
		backgroundFillStyles: fills(scheme.backgroundFillStyles, 'a:bgFillStyleLst'),
	};
}

const UNREAD_FILL: DrawingFill = { kind: 'unsupported', element: 'unread' };

/**
 * The neutral matrices a pptx view was projected from, for the shared resolvers. An entry built
 * by hand without its neutral half reads as unsupported (a fill), empty (a line) or no effects.
 */
export function drawingFormatSchemeOf(scheme: PptxThemeFormatScheme): ThemeFormatScheme {
	const fills = (styles: PptxThemeFillStyle[]) =>
		styles.map((style) => ({ fill: style.fill ?? UNREAD_FILL }));
	return {
		...(scheme.name ? { name: scheme.name } : {}),
		fillStyles: fills(scheme.fillStyles),
		backgroundFillStyles: fills(scheme.backgroundFillStyles),
		lineStyles: scheme.lineStyles.map((style) => style.line ?? {}),
		effectStyles: scheme.effectStyles.map((style) =>
			style.effects ? { effects: style.effects } : {},
		),
	};
}
