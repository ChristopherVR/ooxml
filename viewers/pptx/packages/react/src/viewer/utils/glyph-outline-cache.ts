/**
 * Module-scoped WordArt glyph-outline font cache (see `pptx-viewer-shared`'s
 * `text-warp-outline-font-cache.ts`).
 *
 * A singleton, not React state: `warp-text-renderer.tsx` reads it directly
 * during render (many small `<EnvelopeGlyph>`s deep in the tree), so
 * threading it through props/context for every element would be a much
 * bigger change than the value it buys. `useLoadContent` registers the
 * deck's embedded fonts here before the first render that uses them (see its
 * `setEmbeddedFonts` call site). Google catalogue webfont bytes arrive later
 * (a real fetch); `useFontInjection` bumps its own state once they land so
 * components re-render and pick them up from this same cache. The cache loads
 * its font parser on the first envelope glyph that needs it, and
 * `WarpedText` subscribes to the cache so mounted glyphs repaint once it has.
 */
import { GlyphOutlineFontCache, createGlyphOutlineLookup } from 'ooxml-ui/pptx';

export const glyphOutlineFontCache = /* @__PURE__ */ new GlyphOutlineFontCache();

/** The `getGlyphOutline` callback `buildGlyphEnvelope` accepts, backed by the module singleton. */
export const getGlyphOutline = /* @__PURE__ */ createGlyphOutlineLookup(glyphOutlineFontCache);
