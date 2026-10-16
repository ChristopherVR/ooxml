import type {
	PptxElement,
	PptxSlide,
	PptxThemeColorScheme,
	PptxSlideBackgroundStyleRef,
	XmlObject,
} from '../../types';
import {
	backgroundStyleColorMap,
	backgroundStylePreset,
	isDarkColorMap,
	type PptxBackgroundStylePreset,
} from '../../types/background-style';
import { backgroundRefNode } from '../../utils/background-style-ref';
import { reResolveElementsForColorMapChange } from '../../utils/theme-switching';
import { withLeadingBackground } from './master-save-helpers';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeThemeEffects';

/** Which slides a background style lands on. */
export type BackgroundStyleScope =
	| { kind: 'all' }
	| { kind: 'slides'; slideIds: readonly string[] };

/** Slides and template elements after a background style was applied. */
export interface BackgroundStyleApplication {
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
}

interface ResolvedBackground {
	color: string | undefined;
	gradient: string | undefined;
}

function schemeFromColorMap(map: Readonly<Record<string, string>>): PptxThemeColorScheme {
	const pick = (key: string): string => map[key] ?? '';
	return {
		dk1: pick('dk1'),
		lt1: pick('lt1'),
		dk2: pick('dk2'),
		lt2: pick('lt2'),
		accent1: pick('accent1'),
		accent2: pick('accent2'),
		accent3: pick('accent3'),
		accent4: pick('accent4'),
		accent5: pick('accent5'),
		accent6: pick('accent6'),
		hlink: pick('hlink'),
		folHlink: pick('folHlink'),
	};
}

/**
 * Design > Variants > Background Styles.
 *
 * A style is a `p:bgRef` into the theme's `a:bgFillStyleLst` plus, for the dark columns, a swapped
 * colour map. "Apply to All Slides" writes both on the slide master and makes every slide follow
 * it; "Apply to Selected Slides" writes the same `bgRef` and an explicit `p:clrMapOvr` on each
 * slide (what PowerPoint writes, verified against `Slide.BackgroundStyle`).
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	private resolveBackgroundStyle(preset: PptxBackgroundStylePreset): ResolvedBackground {
		const xml = {
			'p:sld': { 'p:cSld': { 'p:bg': backgroundRefNode(preset.idx, preset.scheme) } },
		} as XmlObject;
		return {
			color: this.extractBackgroundColor(xml),
			gradient: this.extractBackgroundGradient(xml) || undefined,
		};
	}

	private layoutAuthorsBackground(layoutPath: string | undefined): boolean {
		if (!layoutPath) {
			return false;
		}
		const layout = this.layoutXmlMap.get(layoutPath)?.['p:sldLayout'] as XmlObject | undefined;
		return Boolean((layout?.['p:cSld'] as XmlObject | undefined)?.['p:bg']);
	}

	/**
	 * Apply background style `index` (1 to 12) to `slides` and their template elements. Pure with
	 * respect to the handler's state (a preview); {@link commitMasterBackgroundStyle} makes a
	 * deck-wide application permanent.
	 */
	public applyBackgroundStyleToSlides(
		slides: readonly PptxSlide[],
		templateElementsBySlideId: Readonly<Record<string, PptxElement[]>>,
		index: number,
		scope: BackgroundStyleScope,
	): BackgroundStyleApplication | undefined {
		const preset = backgroundStylePreset(index);
		if (!preset) {
			return undefined;
		}
		const targets = scope.kind === 'all' ? undefined : new Set(scope.slideIds);
		const saved = {
			colors: this.themeColorMap,
			fonts: this.themeFontMap,
			format: this.themeFormatScheme,
			clrMap: this.currentMasterClrMap,
			slideClrMap: this.currentSlideClrMapOverride,
		};
		try {
			const templates = { ...templateElementsBySlideId };
			const result = slides.map((slide) => {
				if (targets && !targets.has(slide.id)) {
					return slide;
				}
				const masterPath = slide.layoutPath
					? this.findMasterPathForLayoutBase(slide.layoutPath)
					: undefined;
				this.applyMasterThemeState(masterPath);
				const previousDark = isDarkColorMap(slide.clrMapOverride ?? this.currentMasterClrMap);
				const clrMap = backgroundStyleColorMap(preset.dark);
				this.currentSlideClrMapOverride = clrMap;
				const resolved = this.resolveBackgroundStyle(preset);
				const inherited = scope.kind === 'all';
				const inheritsLayout = this.layoutAuthorsBackground(slide.layoutPath);
				const ref: PptxSlideBackgroundStyleRef = inherited
					? { kind: 'inherit' }
					: { kind: 'preset', idx: preset.idx, scheme: preset.scheme };
				const painted = inherited && inheritsLayout ? undefined : resolved;
				const next: PptxSlide = {
					...slide,
					backgroundStyleRef: {
						...ref,
						...(painted?.color !== undefined ? { resolvedColor: painted.color } : {}),
						...(painted?.gradient ? { resolvedGradient: painted.gradient } : {}),
					},
				};
				delete next.backgroundImage;
				delete next.backgroundImageProperties;
				delete next.backgroundPattern;
				delete next.backgroundShadeToTitle;
				if (painted) {
					next.backgroundColor = painted.color;
					if (painted.gradient) {
						next.backgroundGradient = painted.gradient;
					} else {
						delete next.backgroundGradient;
					}
				}
				if (inherited) {
					delete next.clrMapOverride;
				} else {
					next.clrMapOverride = clrMap;
				}
				const scheme = schemeFromColorMap(this.themeColorMap);
				next.elements = reResolveElementsForColorMapChange(
					slide.elements,
					scheme,
					previousDark,
					preset.dark,
				);
				const template = templates[slide.id];
				if (template) {
					templates[slide.id] = reResolveElementsForColorMapChange(
						template,
						scheme,
						previousDark,
						preset.dark,
					);
				}
				return next;
			});
			return { slides: result, templateElementsBySlideId: templates };
		} finally {
			this.themeColorMap = saved.colors;
			this.themeFontMap = saved.fonts;
			this.themeFormatScheme = saved.format;
			this.currentMasterClrMap = saved.clrMap;
			this.currentSlideClrMapOverride = saved.slideClrMap;
		}
	}

	/**
	 * Write background style `index` onto every slide master (`p:bg` and `p:clrMap`) so a deck-wide
	 * application is saved. Slides follow it through {@link applyBackgroundStyleToSlides}.
	 */
	public async commitMasterBackgroundStyle(index: number): Promise<boolean> {
		const preset = backgroundStylePreset(index);
		if (!preset) {
			return false;
		}
		const masterPaths = new Set<string>([
			...this.masterXmlMap.keys(),
			...this.masterThemePaths.keys(),
		]);
		const clrMap = backgroundStyleColorMap(preset.dark);
		const bg = backgroundRefNode(preset.idx, preset.scheme);
		for (const masterPath of masterPaths) {
			const xml = await this.resolveCachedMasterXml(masterPath);
			const root = xml?.['p:sldMaster'] as XmlObject | undefined;
			if (!root) {
				continue;
			}
			root['p:cSld'] = withLeadingBackground((root['p:cSld'] ?? {}) as XmlObject, bg);
			root['p:clrMap'] = Object.fromEntries(
				Object.entries(clrMap).map(([alias, slot]) => [`@_${alias}`, slot]),
			);
			this.masterClrMaps.set(masterPath, { ...clrMap });
			this.masterStyleOverrides.set(masterPath, { bg, clrMap: { ...clrMap } });
		}
		return true;
	}
}
