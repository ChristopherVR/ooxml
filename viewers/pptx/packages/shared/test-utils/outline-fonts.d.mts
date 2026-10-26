import type { PptxEmbeddedFont, PptxHandler } from 'pptx-viewer-core';

export function embeddedFontList(name?: string): PptxEmbeddedFont[];

export function embedFontsOnLoad(Handler: typeof PptxHandler): PptxEmbeddedFont[][];

/** `cache` is a `GlyphOutlineFontCache`. */
export function retainedFontLists(cache: object): Set<readonly PptxEmbeddedFont[]>;
