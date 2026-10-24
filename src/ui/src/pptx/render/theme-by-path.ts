/**
 * Presentation properties > Theme: point the first slide master, or every one, at another theme
 * part already in the package, in place, for every binding.
 *
 * Reloading the deck after the switch (what React, Vue and Svelte did) wiped Undo and sent the view
 * back to the first slide; re-pointing the masters only (Angular, vanilla) left the slides in the
 * old theme's colours and fonts until the file was reopened. {@link applyThemeByPath} reads the
 * target part, re-points the masters, and re-resolves the colours and fonts of exactly the slides
 * (and their template-element layer) whose layout belongs to a re-targeted master, each against
 * its own master's old part. The binding commits the returned {@link ThemeByPathChange} as ONE
 * undo step; the step records the masters' theme parts (`historyTheme`), so Undo points them back.
 *
 * @module render/theme-by-path
 */
import {
	buildThemeColorMap,
	reResolveElementColors,
	reResolveElementFonts,
	reResolveSlideColors,
	reResolveSlideFonts,
} from 'ooxml-core/pptx';
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

import type { ThemePartSchemes, ThemePartWriter } from './theme-history';
import { queueThemeWrite } from './theme-write-queue';

/** The handler operations the switch needs (the `PptxHandler` has both). */
export type ThemeByPathHandler = Pick<
	ThemePartWriter,
	'getThemePartSchemes' | 'setSlideMasterTheme'
>;

/** A slide master as a binding keeps it (`PptxSlideMaster` fits). */
export interface ThemeByPathMaster {
	readonly path: string;
	readonly themePath?: string | undefined;
	readonly layoutPaths?: readonly string[] | undefined;
	readonly layouts?: readonly { readonly path: string }[] | undefined;
}

/** Any theme-shaped value: a binding's `PptxTheme`, or the looser record some bindings keep. */
interface ThemeLike {
	colorScheme?: unknown;
	fontScheme?: unknown;
	name?: unknown;
}

export interface ThemeByPathInput<M extends ThemeByPathMaster, T extends ThemeLike> {
	/** The theme part to switch to (`PptxThemeOption.path`). */
	readonly themePath: string;
	/** Every master, or only the first one (the inspector's two buttons). */
	readonly applyToAllMasters: boolean;
	readonly slides: readonly PptxSlide[];
	/** The binding's masters, in deck order: the first one's part is the deck's theme. */
	readonly slideMasters: readonly M[];
	/** The binding's deck theme (the first master's). */
	readonly theme: T | null | undefined;
	/** The master/layout element layer, for bindings that keep it apart from the slides. */
	readonly templateElementsBySlideId?: Readonly<Record<string, PptxElement[]>> | undefined;
}

/** What a binding commits, as one undo step. */
export interface ThemeByPathChange<M extends ThemeByPathMaster, T extends ThemeLike> {
	/** Every slide; those on a master that kept its part are the same objects. */
	readonly slides: PptxSlide[];
	/** The template layer re-resolved the same way (empty when the input had none). */
	readonly templateElementsBySlideId: Record<string, PptxElement[]>;
	/** The masters, re-targeted ones on the new part. */
	readonly slideMasters: M[];
	/**
	 * When the first master was re-targeted, the deck theme carrying the new part's schemes and name,
	 * and the colour map built from it. Absent when the deck theme is unchanged.
	 */
	readonly deckTheme?: { readonly theme: T; readonly themeColorMap: Record<string, string> };
}

/** The layouts a master owns. */
function layoutsOf(master: ThemeByPathMaster): readonly string[] {
	return master.layoutPaths ?? master.layouts?.map((layout) => layout.path) ?? [];
}

/** `theme` carrying the schemes and name of `part` (no name when the part has none). */
function withPartTheme<T extends ThemeLike>(
	theme: T | null | undefined,
	part: ThemePartSchemes,
): T {
	const next: Record<string, unknown> = {
		...theme,
		colorScheme: part.colorScheme,
		fontScheme: part.fontScheme,
	};
	if (part.name) {
		next.name = part.name;
	} else {
		delete next.name;
	}
	return next as T;
}

/** Re-resolve one element layer from the `from` part's schemes to the `to` part's. */
function reResolveElements(
	elements: PptxElement[],
	from: ThemePartSchemes,
	to: ThemePartSchemes,
): PptxElement[] {
	const colored = reResolveElementColors(
		elements,
		buildThemeColorMap(from.colorScheme),
		to.colorScheme,
	);
	return reResolveElementFonts(colored, from.fontScheme, to.fontScheme);
}

/** Re-resolve one slide from the `from` part's schemes to the `to` part's. */
function reResolveSlide(slide: PptxSlide, from: ThemePartSchemes, to: ThemePartSchemes): PptxSlide {
	const [colored] = reResolveSlideColors(
		[slide],
		buildThemeColorMap(from.colorScheme),
		to.colorScheme,
	);
	const [themed] = reResolveSlideFonts([colored ?? slide], from.fontScheme, to.fontScheme);
	return themed ?? slide;
}

/**
 * Point the masters at `themePath` in `handler` and return what the binding commits, or
 * `undefined` when there is nothing to do (no handler, no such part, or every target master is on
 * it already). Runs on the handler's theme write queue: it waits for the theme writes queued
 * before it, and a save waits for it.
 */
export function applyThemeByPath<M extends ThemeByPathMaster, T extends ThemeLike>(
	handler: ThemeByPathHandler | null | undefined,
	input: ThemeByPathInput<M, T>,
): Promise<ThemeByPathChange<M, T> | undefined> {
	const { themePath, slideMasters } = input;
	const candidates = input.applyToAllMasters ? slideMasters : slideMasters.slice(0, 1);
	const targets = candidates.filter((master) => master.themePath !== themePath);
	if (!handler || targets.length === 0) {
		return Promise.resolve(undefined);
	}
	return queueThemeWrite(handler, async () => {
		const next = await handler.getThemePartSchemes(themePath);
		if (!next) {
			return undefined;
		}
		// Each re-targeted master's slides were resolved against its own old part.
		const previous = new Map<string, ThemePartSchemes | undefined>();
		for (const master of targets) {
			previous.set(
				master.path,
				master.themePath ? await handler.getThemePartSchemes(master.themePath) : undefined,
			);
			await handler.setSlideMasterTheme(master.path, themePath);
		}
		const masterOfLayout = new Map<string, string>();
		for (const master of slideMasters) {
			for (const layout of layoutsOf(master)) {
				masterOfLayout.set(layout, master.path);
			}
		}
		// A slide whose layout no master lists draws with the deck theme, the first master's.
		const fromOf = (slide: PptxSlide): ThemePartSchemes | undefined => {
			const layoutMaster = slide.layoutPath ? masterOfLayout.get(slide.layoutPath) : undefined;
			const masterPath = layoutMaster ?? slideMasters[0]?.path;
			return masterPath ? previous.get(masterPath) : undefined;
		};
		const templates = input.templateElementsBySlideId ?? {};
		const templateElementsBySlideId: Record<string, PptxElement[]> = { ...templates };
		const slides = input.slides.map((slide) => {
			const from = fromOf(slide);
			if (!from) {
				return slide;
			}
			const layer = templates[slide.id];
			if (layer) {
				templateElementsBySlideId[slide.id] = reResolveElements(layer, from, next);
			}
			return reResolveSlide(slide, from, next);
		});
		const first = slideMasters[0];
		const change: ThemeByPathChange<M, T> = {
			slides,
			templateElementsBySlideId,
			slideMasters: slideMasters.map((master) =>
				previous.has(master.path) ? { ...master, themePath } : master,
			),
		};
		if (!first || !previous.has(first.path)) {
			return change;
		}
		const theme = withPartTheme(input.theme, next);
		return { ...change, deckTheme: { theme, themeColorMap: buildThemeColorMap(next.colorScheme) } };
	});
}
