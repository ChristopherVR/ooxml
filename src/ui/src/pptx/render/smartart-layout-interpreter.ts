/**
 * SmartArt DiagramML interpreter - public entry point + dispatch.
 *
 * Re-exported from `ooxml-core/pptx`, which is now the single home of this
 * interpreter: `ooxml-core/pptx`'s save/decompose pipeline fabricates the
 * cached `dsp:` diagram drawing using the SAME interpreter this package's SVG
 * -fallback preview path calls, so the fabricated drawing on save matches
 * what every binding renders on screen (`composite`/`conn`/`sp`/`tx`,
 * decided `dgm:choose`/`dgm:forEach`, and manual `cust*` node overrides all
 * included). `ooxml-core/pptx` cannot import `pptx-viewer-shared` (this
 * package depends on core, not the other way around, and core is published
 * standalone), so this is the only direction that avoids a circular
 * dependency.
 */

export { interpretSmartArtLayout, type InterpretLayoutInput } from 'ooxml-core/pptx';
