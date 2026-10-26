import { createGlyphOutlineLookup, GlyphOutlineFontCache } from 'ooxml-ui/pptx';

/**
 * Module-scoped WordArt glyph-outline font cache (see `pptx-viewer-shared`'s
 * `text-warp-outline-font-cache.ts`). Svelte port of the React/Vue/Angular
 * bindings' `glyph-outline-cache`; same rationale: a module singleton, not
 * component state, since `WordArtText.svelte` reads it directly during
 * render for every glyph.
 *
 * `presentation-loader.svelte.ts` registers the deck's embedded fonts here
 * SYNCHRONOUSLY (in its `load()` method, alongside `this.embeddedFonts = ...`,
 * so before the effects that would render them run). `viewer-effects.svelte.ts`
 * registers catalogue webfont bytes once fetched and bumps
 * {@link glyphOutlineFontsTick}'s `value` so `WordArtText.svelte`'s `$derived`
 * glyph layout (which reads it) recomputes and picks them up from this same
 * cache. The cache loads its font parser on the first envelope glyph that
 * needs it and notifies its subscribers when that fills in outlines, which
 * bumps the same tick.
 */
export const glyphOutlineFontCache = new GlyphOutlineFontCache();

/** A `$state` box so a plain module export stays reactive (runes need a `.svelte.ts` file). */
class GlyphOutlineFontsTick {
	value = $state(0);
}

/** Bumped whenever a webfont's outline bytes are newly registered; read (not written) by `WordArtText.svelte`. */
export const glyphOutlineFontsTick = new GlyphOutlineFontsTick();
glyphOutlineFontCache.subscribe(() => {
	glyphOutlineFontsTick.value += 1;
});

/** The `getGlyphOutline` callback `buildGlyphEnvelope` accepts, backed by the module singleton. */
export const getGlyphOutline = createGlyphOutlineLookup(glyphOutlineFontCache);
