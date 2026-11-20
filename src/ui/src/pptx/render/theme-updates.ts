/**
 * A theme change applied to the deck state: the slides and the master/layout elements re-resolved
 * against the new colour and font schemes, and the theme and colour map that go with them.
 *
 * Synchronous and handler-free, so a binding can take the result into its state (and its undo
 * history) at once and write the theme parts afterwards (`writeThemeParts`). The Design galleries
 * and the assistant's theme edits both go through it.
 *
 * @module render/theme-updates
 */
import type {
	PptxElement,
	PptxTheme,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
} from 'ooxml-core/pptx';
import {
	buildThemeColorMap,
	reResolveElementColors,
	reResolveElementFonts,
	reResolveSlideColors,
	reResolveSlideFonts,
} from 'ooxml-core/pptx';

import type { DeckStyleState } from './deck-style';

type Templates = Record<string, PptxElement[]>;

function mapTemplates(
	templates: Templates,
	map: (elements: PptxElement[]) => PptxElement[],
): Templates {
	return Object.fromEntries(Object.entries(templates).map(([id, elements]) => [id, map(elements)]));
}

export function withColors(
	state: DeckStyleState,
	colorScheme: PptxThemeColorScheme,
): DeckStyleState {
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

export function withFonts(state: DeckStyleState, fontScheme: PptxThemeFontScheme): DeckStyleState {
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

export function withName(state: DeckStyleState, name: string | undefined): DeckStyleState {
	return name ? { ...state, theme: { ...state.theme, name } } : state;
}

/** The state partial theme `updates` (colour scheme, font scheme, name) produce from `base`. */
export function withThemeUpdates(
	base: DeckStyleState,
	updates: Partial<Pick<PptxTheme, 'colorScheme' | 'fontScheme' | 'name'>>,
): DeckStyleState {
	const colored = updates.colorScheme ? withColors(base, updates.colorScheme) : base;
	const fonted = updates.fontScheme ? withFonts(colored, updates.fontScheme) : colored;
	return withName(fonted, updates.name);
}
