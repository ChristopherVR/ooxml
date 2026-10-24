import type {
	PptxElement,
	PptxSlide,
	PptxThemeFormatScheme,
	PptxThemeVariant,
	ResolvedStyleMatrix,
	ShapeStyle,
	XmlObject,
} from '../../types';
import {
	findThemeEffectScheme,
	wrapFormatSchemeXml,
	type PptxThemeEffectScheme,
} from '../../utils/theme-effect-schemes';
import {
	STYLE_MATRIX_EFFECT_KEYS,
	STYLE_MATRIX_FILL_KEYS,
	STYLE_MATRIX_LINE_KEYS,
	captureStyleBaseline,
	effectIsPurelyStyleMatrix,
	fillIsPurelyStyleMatrix,
} from './authored-shape-style';
import { HeldFormatSchemes } from './held-format-schemes';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeStyleMatrixResolve';

/** Slides and template (master/layout) elements after style references were resolved again. */
export interface StyleReferenceReResolution {
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
}

/** Copy `keys` of `from` onto a copy of `style`, dropping the keys `from` leaves unset. */
function withKeys(
	style: ShapeStyle,
	from: ShapeStyle,
	keys: readonly (keyof ShapeStyle)[],
): ShapeStyle {
	const next: Record<string, unknown> = { ...style };
	for (const key of keys) {
		if (from[key] === undefined) {
			delete next[key];
		} else {
			next[key] = from[key];
		}
	}
	return next as ShapeStyle;
}

function refNode(idx: number, colorXml: XmlObject | undefined): XmlObject {
	return { '@_idx': String(idx), ...(colorXml ?? {}) };
}

/**
 * Design > Variants > Effects: replace the theme's format scheme (`a:fmtScheme`) and redraw the
 * shapes that reference it.
 *
 * A shape keeps `fillRef`/`lnRef`/`effectRef` indices and the flat values they resolved to at load
 * time. {@link reResolveThemeStyleReferences} resolves the indices again against a format scheme
 * and replaces only what the shape still inherits (its flat value equals the recorded baseline),
 * so a fill, outline or shadow the author set by hand survives a scheme change, as in PowerPoint.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	private readonly effectSchemeCache = new Map<string, PptxThemeFormatScheme>();

	/** Every `a:fmtScheme` written into a theme part or over one ({@link restoreThemeFormatScheme}). */
	private readonly heldFormatSchemes = new HeldFormatSchemes();

	/**
	 * Resolve a `<p:style>` node as the base class does, against built-in effect scheme
	 * `effectSchemeId` instead of the deck's own (the Effects gallery draws each tile this way).
	 */
	public override resolveStyleMatrixReferences(
		styleXml: XmlObject,
		effectSchemeId?: string,
	): ResolvedStyleMatrix {
		const preset = effectSchemeId ? findThemeEffectScheme(effectSchemeId) : undefined;
		if (!preset) {
			return super.resolveStyleMatrixReferences(styleXml);
		}
		let scheme = this.effectSchemeCache.get(preset.id);
		if (!scheme) {
			scheme = this.parseEffectScheme(preset)?.scheme;
			if (scheme) {
				this.effectSchemeCache.set(preset.id, scheme);
			}
		}
		const saved = this.themeFormatScheme;
		if (scheme) {
			this.themeFormatScheme = scheme;
		}
		try {
			return super.resolveStyleMatrixReferences(styleXml);
		} finally {
			this.themeFormatScheme = saved;
		}
	}

	private parseEffectScheme(
		preset: Pick<PptxThemeEffectScheme, 'fmtSchemeXml'>,
	): { node: XmlObject; scheme: PptxThemeFormatScheme } | undefined {
		const parsed = this.parser.parse(wrapFormatSchemeXml(preset.fmtSchemeXml)) as XmlObject;
		const node = parsed['a:fmtScheme'] as XmlObject | undefined;
		if (!node) {
			return undefined;
		}
		delete node['@_xmlns:a'];
		return { node, scheme: this.parseFormatScheme(node) };
	}

	/**
	 * Write a built-in effect scheme into every theme part a master references and refresh the
	 * in-memory format schemes. Slides are not touched: call {@link reResolveThemeStyleReferences}
	 * for them.
	 *
	 * @returns the parsed scheme now in force, or `undefined` for an unknown id
	 */
	public async applyThemeEffectScheme(id: string): Promise<PptxThemeFormatScheme | undefined> {
		const preset = findThemeEffectScheme(id);
		return preset ? this.applyThemeFormatSchemeXml(preset.fmtSchemeXml) : undefined;
	}

	/**
	 * Write an `a:fmtScheme` element (as the text a theme part carries, `a` prefix undeclared)
	 * into every theme part a master references; a real theme variant brings its own.
	 */
	public async applyThemeFormatSchemeXml(
		fmtSchemeXml: string,
	): Promise<PptxThemeFormatScheme | undefined> {
		const parsed = this.parseEffectScheme({ fmtSchemeXml });
		return parsed ? this.writeFormatScheme(() => parsed.node) : undefined;
	}

	/**
	 * Put back a format scheme the deck held before (Undo and Redo of an Effects pick), as the parts
	 * held it: each theme part a master references gets back what it held the last time the deck's
	 * scheme was `scheme` over those same parts, so a part with a scheme of its own gets its own back
	 * (with no such time, `scheme`'s own element goes into every one). With `themePaths`, `scheme`'s
	 * own element goes into those parts only.
	 *
	 * @returns the deck's format scheme now, or `undefined` (nothing written) for a scheme no part
	 *   held while this handler wrote into it
	 */
	public async restoreThemeFormatScheme(
		scheme: PptxThemeFormatScheme,
		themePaths?: readonly string[],
	): Promise<PptxThemeFormatScheme | undefined> {
		const parts = themePaths
			? undefined
			: this.heldFormatSchemes.parts(scheme, await this.resolveThemeEditTargets());
		if (parts) {
			return this.writeFormatScheme((path) => parts.get(path));
		}
		const element = this.heldFormatSchemes.element(scheme);
		return element ? this.writeFormatScheme(() => element, themePaths) : undefined;
	}

	/**
	 * Write `elementFor(part)` into every theme part a master references (or only `themePaths`),
	 * skipping a part it gives none, holding what each part held and now holds, and refresh the
	 * in-memory format schemes of the masters on those parts and, when the deck's part is among
	 * them, the deck's.
	 */
	private async writeFormatScheme(
		elementFor: (themePath: string) => XmlObject | undefined,
		themePaths?: readonly string[],
	): Promise<PptxThemeFormatScheme | undefined> {
		const deckBefore = this.getThemeFormatScheme();
		const replaced = new Map<string, XmlObject>();
		const schemes = new Map<string, PptxThemeFormatScheme>();
		const written = await this.rewriteThemeParts(themePaths, (data, path) => {
			const element = elementFor(path);
			const elements = (data['a:theme'] as XmlObject | undefined)?.['a:themeElements'] as
				| XmlObject
				| undefined;
			if (!element || !elements) {
				return false;
			}
			const previous = elements['a:fmtScheme'] as XmlObject | undefined;
			if (previous) {
				this.heldFormatSchemes.hold(this.parseFormatScheme(previous), previous);
				replaced.set(path, structuredClone(previous));
			}
			const scheme = this.parseFormatScheme(element);
			this.heldFormatSchemes.hold(scheme, element);
			elements['a:fmtScheme'] = structuredClone(element);
			schemes.set(path, scheme);
			return true;
		});
		if (written.length === 0) {
			return undefined;
		}
		if (deckBefore) {
			this.heldFormatSchemes.holdParts(deckBefore, replaced);
		}
		for (const path of written) {
			const xml = await this.zip.file(path)?.async('string');
			if (xml !== undefined) {
				this.originalThemeXmlByPath.set(path, xml);
			}
			for (const masterPath of this.mastersUsingThemes([path])) {
				this.masterThemeFormatSchemes.set(masterPath, schemes.get(path)!);
			}
		}
		const primary = await this.resolvePrimaryThemePath();
		const deck = primary === undefined ? undefined : schemes.get(primary);
		if (deck) {
			this.globalThemeFormatSchemeSnapshot = deck;
			this.themeFormatScheme = deck;
			this.heldFormatSchemes.holdParts(
				deck,
				new Map(written.map((path) => [path, elementFor(path)!])),
			);
		}
		return this.getThemeFormatScheme();
	}

	/** The real theme variants the loaded package carries, or none. */
	public getThemeVariants(): PptxThemeVariant[] {
		return this.loadedThemeVariants;
	}

	/** The format scheme the deck's primary theme carries now. */
	public getThemeFormatScheme(): PptxThemeFormatScheme | undefined {
		return this.globalThemeFormatSchemeSnapshot ?? this.themeFormatScheme;
	}

	/**
	 * Resolve every shape's `p:style` references again.
	 *
	 * @param effectSchemeId resolve against this built-in scheme instead of the one in force
	 *   (a preview); omit it after {@link applyThemeEffectScheme}
	 */
	public reResolveThemeStyleReferences(
		slides: readonly PptxSlide[],
		templateElementsBySlideId: Readonly<Record<string, PptxElement[]>>,
		effectSchemeId?: string,
	): StyleReferenceReResolution {
		const preset = effectSchemeId ? findThemeEffectScheme(effectSchemeId) : undefined;
		const override = preset ? this.parseEffectScheme(preset)?.scheme : undefined;
		const saved = {
			colors: this.themeColorMap,
			fonts: this.themeFontMap,
			format: this.themeFormatScheme,
			clrMap: this.currentMasterClrMap,
			slideClrMap: this.currentSlideClrMapOverride,
		};
		try {
			const templates: Record<string, PptxElement[]> = {};
			const resolved = slides.map((slide) => {
				this.applyMasterThemeState(
					slide.layoutPath ? this.findMasterPathForLayoutBase(slide.layoutPath) : undefined,
				);
				this.currentSlideClrMapOverride = slide.clrMapOverride ?? null;
				if (override) {
					this.themeFormatScheme = override;
				}
				const template = templateElementsBySlideId[slide.id];
				if (template) {
					templates[slide.id] = template.map((element) => this.reResolveElementStyle(element));
				}
				return { ...slide, elements: slide.elements.map((el) => this.reResolveElementStyle(el)) };
			});
			for (const [slideId, elements] of Object.entries(templateElementsBySlideId)) {
				templates[slideId] ??= elements;
			}
			return { slides: resolved, templateElementsBySlideId: templates };
		} finally {
			this.themeColorMap = saved.colors;
			this.themeFontMap = saved.fonts;
			this.themeFormatScheme = saved.format;
			this.currentMasterClrMap = saved.clrMap;
			this.currentSlideClrMapOverride = saved.slideClrMap;
		}
	}

	private reResolveElementStyle(element: PptxElement): PptxElement {
		const withChildren =
			element.type === 'group'
				? {
						...element,
						children: element.children.map((child) => this.reResolveElementStyle(child)),
					}
				: element;
		const style = 'shapeStyle' in withChildren ? withChildren.shapeStyle : undefined;
		if (!style) {
			return withChildren;
		}
		const next = this.reResolveShapeStyle(style);
		return next === style ? withChildren : ({ ...withChildren, shapeStyle: next } as PptxElement);
	}

	private reResolveShapeStyle(style: ShapeStyle): ShapeStyle {
		const styleXml: XmlObject = {};
		if (style.fillRefIdx !== undefined) {
			styleXml['a:fillRef'] = refNode(style.fillRefIdx, style.fillRefColorXml);
		}
		if (style.lnRefIdx !== undefined) {
			styleXml['a:lnRef'] = refNode(style.lnRefIdx, style.lnRefColorXml);
		}
		if (style.effectRefIdx !== undefined) {
			styleXml['a:effectRef'] = refNode(style.effectRefIdx, style.effectRefColorXml);
		}
		if (Object.keys(styleXml).length === 0) {
			return style;
		}
		const fresh = this.resolveStyleMatrixReferences(styleXml).shapeStyle;
		let next = style;
		if (style.fillRefIdx !== undefined && fillIsPurelyStyleMatrix(style)) {
			next = {
				...withKeys(next, fresh, STYLE_MATRIX_FILL_KEYS),
				inheritedFillStyle: captureStyleBaseline(fresh, STYLE_MATRIX_FILL_KEYS),
			};
		}
		const lineBaseline = style.inheritedLineStyle;
		if (style.lnRefIdx !== undefined && lineBaseline) {
			const inherited = STYLE_MATRIX_LINE_KEYS.filter((key) => style[key] === lineBaseline[key]);
			next = {
				...withKeys(next, fresh, inherited),
				inheritedLineStyle: captureStyleBaseline(fresh, STYLE_MATRIX_LINE_KEYS),
			};
		}
		if (style.effectRefIdx !== undefined && effectIsPurelyStyleMatrix(style)) {
			next = {
				...withKeys(next, fresh, STYLE_MATRIX_EFFECT_KEYS),
				inheritedEffectStyle: captureStyleBaseline(fresh, STYLE_MATRIX_EFFECT_KEYS),
			};
		}
		return next;
	}
}
