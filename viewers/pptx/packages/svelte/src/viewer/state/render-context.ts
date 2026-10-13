import type { CanvasSize, TableStyleContext } from 'ooxml-ui/pptx';
import type {
	ParsedTableStyleMap,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemeFormatScheme,
	PptxTheme,
} from 'pptx-viewer-core';
import { getContext, setContext } from 'svelte';

const RENDER_CONTEXT = Symbol('pptx-svelte-render-context');

export interface RenderContextSource {
	getColorScheme: () => PptxThemeColorScheme | undefined;
	getTableStyleMap: () => ParsedTableStyleMap | undefined;
	/** Theme font scheme, so band/header cell text can resolve major/minor fonts. */
	getFontScheme?: () => PptxThemeFontScheme | undefined;
	/** Theme format scheme, so a style's `a:tblBg/a:fillRef` and `a:lnRef` borders resolve. */
	getFormatScheme?: () => PptxThemeFormatScheme | undefined;
	/** Deck canvas size, so template insertion can target the real slide surface. */
	getCanvasSize?: () => CanvasSize | undefined;
}

/** What {@link renderContextSourceFor} reads from the deck loader. */
export interface RenderContextLoader {
	readonly colorScheme: PptxThemeColorScheme | undefined;
	readonly tableStyleMap: ParsedTableStyleMap | undefined;
	readonly presentationTheme: PptxTheme | undefined;
	readonly canvasSize: CanvasSize | undefined;
}

/** The render-context source over a deck loader; every getter reads it live. */
export function renderContextSourceFor(loader: RenderContextLoader): RenderContextSource {
	return {
		getColorScheme: () => loader.colorScheme,
		getTableStyleMap: () => loader.tableStyleMap,
		getFontScheme: () => loader.presentationTheme?.fontScheme,
		getFormatScheme: () => loader.presentationTheme?.formatScheme,
		getCanvasSize: () => loader.canvasSize,
	};
}

export function provideRenderContext(source: RenderContextSource): void {
	setContext(RENDER_CONTEXT, source);
}

/**
 * The raw render-context source, for chrome that needs the deck's colour
 * scheme or canvas size outside a table (e.g. the Slide Templates gallery).
 * Must be called during component init; invoke the getters lazily so reads
 * stay live against the loader's runes state.
 */
export function getRenderContextSource(): RenderContextSource | undefined {
	return getContext<RenderContextSource | undefined>(RENDER_CONTEXT);
}

export function useTableStyleContext(): TableStyleContext | undefined {
	return tableStyleContextFrom(getContext<RenderContextSource | undefined>(RENDER_CONTEXT));
}

/**
 * The table-style context a render-context source supplies (style map plus the
 * theme's colour, font and format schemes), or `undefined` when it has none.
 */
export function tableStyleContextFrom(
	source: RenderContextSource | undefined,
): TableStyleContext | undefined {
	if (!source) {
		return undefined;
	}
	const colorScheme = source.getColorScheme();
	const tableStyleMap = source.getTableStyleMap();
	const fontScheme = source.getFontScheme?.();
	const formatScheme = source.getFormatScheme?.();
	return colorScheme || tableStyleMap || fontScheme || formatScheme
		? { colorScheme, tableStyleMap, fontScheme, formatScheme }
		: undefined;
}
