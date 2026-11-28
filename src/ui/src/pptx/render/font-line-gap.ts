/**
 * The typographic line gap PowerPoint adds above the ascent of fonts that ask
 * for their typo metrics, which a browser does not expose.
 *
 * Measured on PowerPoint 16 exports of an `H` at 72pt, single spaced, in every
 * font in the Windows font folder (235 families). For a font without the
 * `USE_TYPO_METRICS` bit (`OS/2.fsSelection` bit 7) the baseline is the Windows
 * ascent scaled to the 1.2 em pitch, whatever its line gap (Calibri, Arial and
 * Consolas have large gaps and are within 1px of that). A font with the bit
 * uses `sTypoAscender` and `sTypoDescender`, and `sTypoLineGap` is added above
 * the ascent before the box is scaled to the pitch: Gabriola (gap 0.70 em)
 * puts its baseline at 0.977 em, where ascent and descent alone say 0.82 em.
 * Canvas `fontBoundingBox*` metrics already follow the same typo ascent and
 * descent but leave the gap out.
 *
 * The gaps live in `ooxml-core/text` (`typoLineGapEm`), which Word's canvas
 * measurer reads too; this module keeps the renderer's import path.
 *
 * @module render/font-line-gap
 */
export { typoLineGapEm } from 'ooxml-core/text';
