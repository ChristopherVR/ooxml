/**
 * Design galleries applied to the deck state, and the hover preview of the same results.
 *
 * Every binding holds the same four pieces of state a Design gallery changes: the slides, the
 * master/layout elements drawn behind them, the theme and its colour map. {@link previewGalleryResult}
 * computes the state a result would produce without touching the handler, {@link commitGalleryResult}
 * writes it to the handler too, and {@link createGalleryPreviewController} shows a preview over
 * the live state and puts the original back, never recording history.
 *
 * @module render/deck-style
 */
import type {
	BackgroundStyleApplication,
	BackgroundStyleScope,
	PptxElement,
	PptxSlide,
	PptxTheme,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemeFormatScheme,
	StyleReferenceReResolution,
} from 'ooxml-core/pptx';
import {
	buildThemeColorMap,
	reResolveElementColors,
	reResolveElementFonts,
	reResolveSlideColors,
	reResolveSlideFonts,
} from 'ooxml-core/pptx';

import type { RibbonGalleryApplyResult } from './ribbon-galleries/gallery-types';

/** The state a Design gallery changes. */
export interface DeckStyleState {
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
	theme: PptxTheme | undefined;
	themeColorMap: Record<string, string> | undefined;
}

/** Where a "selected slides" background style lands. */
export interface DeckStyleTarget {
	/** Ids of the slides the thumbnail panel has selected (the active slide when none). */
	selectedSlideIds: readonly string[];
}

/** The handler operations the Design galleries need (the `PptxHandler` has all of them). */
export interface DeckStyleHandler {
	applyTheme(
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme,
		themeName?: string,
	): Promise<void>;
	updateThemeColorScheme(colorScheme: PptxThemeColorScheme): Promise<void>;
	updateThemeFontScheme(fontScheme: PptxThemeFontScheme): Promise<void>;
	applyThemeEffectScheme(id: string): Promise<PptxThemeFormatScheme | undefined>;
	applyThemeFormatSchemeXml(fmtSchemeXml: string): Promise<PptxThemeFormatScheme | undefined>;
	reResolveThemeStyleReferences(
		slides: readonly PptxSlide[],
		templateElementsBySlideId: Readonly<Record<string, PptxElement[]>>,
		effectSchemeId?: string,
	): StyleReferenceReResolution;
	applyBackgroundStyleToSlides(
		slides: readonly PptxSlide[],
		templateElementsBySlideId: Readonly<Record<string, PptxElement[]>>,
		index: number,
		scope: BackgroundStyleScope,
	): BackgroundStyleApplication | undefined;
	commitMasterBackgroundStyle(index: number): Promise<boolean>;
}

type Templates = Record<string, PptxElement[]>;

function mapTemplates(
	templates: Templates,
	map: (elements: PptxElement[]) => PptxElement[],
): Templates {
	return Object.fromEntries(Object.entries(templates).map(([id, elements]) => [id, map(elements)]));
}

function withColors(state: DeckStyleState, colorScheme: PptxThemeColorScheme): DeckStyleState {
	const previous = state.theme?.colorScheme ? buildThemeColorMap(state.theme.colorScheme) : {};
	return {
		...state,
		slides: reResolveSlideColors(state.slides, previous, colorScheme),
		templateElementsBySlideId: mapTemplates(state.templateElementsBySlideId, (elements) =>
			reResolveElementColors(elements, previous, colorScheme),
		),
		theme: { ...state.theme, colorScheme },
		themeColorMap: buildThemeColorMap(colorScheme),
	};
}

function withFonts(state: DeckStyleState, fontScheme: PptxThemeFontScheme): DeckStyleState {
	const previous = state.theme?.fontScheme;
	return {
		...state,
		slides: reResolveSlideFonts(state.slides, previous, fontScheme),
		templateElementsBySlideId: mapTemplates(state.templateElementsBySlideId, (elements) =>
			reResolveElementFonts(elements, previous, fontScheme),
		),
		theme: { ...state.theme, fontScheme },
	};
}

function withName(state: DeckStyleState, name: string | undefined): DeckStyleState {
	return name ? { ...state, theme: { ...state.theme, name } } : state;
}

function backgroundScope(scope: 'all' | 'selected', target: DeckStyleTarget): BackgroundStyleScope {
	return scope === 'all' ? { kind: 'all' } : { kind: 'slides', slideIds: target.selectedSlideIds };
}

/**
 * The state `result` produces from `base`, computed without changing the handler (a preview), or
 * `undefined` when the result changes nothing a preview can show (an element patch, a command).
 */
export function previewGalleryResult(
	result: RibbonGalleryApplyResult,
	base: DeckStyleState,
	handler: DeckStyleHandler | null | undefined,
	target: DeckStyleTarget,
): DeckStyleState | undefined {
	switch (result.kind) {
		case 'theme':
			return withName(
				withFonts(withColors(base, result.colorScheme), result.fontScheme),
				result.name,
			);
		case 'themeColorScheme':
			return withColors(base, result.colorScheme);
		case 'themeFontScheme':
			return withFonts(base, result.fontScheme);
		case 'themeEffects': {
			if (!handler || !result.schemeId) {
				return undefined;
			}
			const out = handler.reResolveThemeStyleReferences(
				base.slides,
				base.templateElementsBySlideId,
				result.schemeId,
			);
			return {
				...base,
				slides: out.slides,
				templateElementsBySlideId: out.templateElementsBySlideId,
			};
		}
		case 'backgroundStyle': {
			const out = handler?.applyBackgroundStyleToSlides(
				base.slides,
				base.templateElementsBySlideId,
				result.styleIndex,
				backgroundScope(result.scope, target),
			);
			return out && { ...base, ...out };
		}
		case 'sequence': {
			let state: DeckStyleState | undefined;
			for (const step of result.steps) {
				state = previewGalleryResult(step, state ?? base, handler, target) ?? state;
			}
			return state;
		}
		default:
			return undefined;
	}
}

/**
 * Write `result` to the handler (the theme part, the master) and return the state it produces
 * from `base`, or `undefined` when there is nothing for the deck state to take.
 */
export async function commitGalleryResult(
	result: RibbonGalleryApplyResult,
	base: DeckStyleState,
	handler: DeckStyleHandler | null | undefined,
	target: DeckStyleTarget,
): Promise<DeckStyleState | undefined> {
	if (!handler) {
		return undefined;
	}
	switch (result.kind) {
		case 'theme': {
			await handler.applyTheme(result.colorScheme, result.fontScheme, result.name);
			return previewGalleryResult(result, base, handler, target);
		}
		case 'themeColorScheme':
			await handler.updateThemeColorScheme(result.colorScheme);
			return previewGalleryResult(result, base, handler, target);
		case 'themeFontScheme':
			await handler.updateThemeFontScheme(result.fontScheme);
			return previewGalleryResult(result, base, handler, target);
		case 'themeEffects': {
			const scheme = result.schemeId
				? await handler.applyThemeEffectScheme(result.schemeId)
				: result.fmtSchemeXml
					? await handler.applyThemeFormatSchemeXml(result.fmtSchemeXml)
					: undefined;
			if (!scheme) {
				return undefined;
			}
			const out = handler.reResolveThemeStyleReferences(
				base.slides,
				base.templateElementsBySlideId,
			);
			return {
				...base,
				slides: out.slides,
				templateElementsBySlideId: out.templateElementsBySlideId,
				theme: { ...base.theme, formatScheme: scheme },
			};
		}
		case 'backgroundStyle': {
			const out = handler.applyBackgroundStyleToSlides(
				base.slides,
				base.templateElementsBySlideId,
				result.styleIndex,
				backgroundScope(result.scope, target),
			);
			if (!out) {
				return undefined;
			}
			if (result.scope === 'all') {
				await handler.commitMasterBackgroundStyle(result.styleIndex);
			}
			return { ...base, ...out };
		}
		case 'sequence': {
			let state: DeckStyleState | undefined;
			for (const step of result.steps) {
				state = (await commitGalleryResult(step, state ?? base, handler, target)) ?? state;
			}
			return state;
		}
		default:
			return undefined;
	}
}

/** What a binding gives the preview controller. */
export interface GalleryPreviewAdapter {
	/** The live state. */
	read(): DeckStyleState;
	/**
	 * Show `state` over the live state without recording an undo step, marking the deck dirty,
	 * autosaving or sending it to collaborators.
	 */
	show(state: DeckStyleState): void;
	/** `true` before the first `show` of a preview, `false` after the original is back. */
	hold?(on: boolean): void;
}

export interface GalleryPreviewController {
	/** Show `state` (computed from {@link GalleryPreviewController.base}); `undefined` ends the preview. */
	preview(state: DeckStyleState | undefined): void;
	/** Put the original back. */
	end(): void;
	/**
	 * Stop previewing before a commit and return the state the commit starts from: the original
	 * when a preview was showing, else `undefined` (read the live state).
	 */
	settle(): DeckStyleState | undefined;
	/** The state a preview is computed from: the original while one shows, else the live state. */
	base(): DeckStyleState;
	readonly active: boolean;
}

/** Show previews over the live state and restore it afterwards. */
export function createGalleryPreviewController(
	adapter: GalleryPreviewAdapter,
): GalleryPreviewController {
	let original: DeckStyleState | undefined;
	return {
		preview(state) {
			if (!state) {
				this.end();
				return;
			}
			if (!original) {
				original = adapter.read();
				adapter.hold?.(true);
			}
			adapter.show(state);
		},
		end() {
			if (!original) {
				return;
			}
			const restore = original;
			original = undefined;
			adapter.show(restore);
			adapter.hold?.(false);
		},
		settle() {
			const was = original;
			this.end();
			return was;
		},
		base() {
			return original ?? adapter.read();
		},
		get active() {
			return original !== undefined;
		},
	};
}
