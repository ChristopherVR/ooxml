/**
 * Webfonts for document fonts the reader has neither installed nor embedded: the bundled Google
 * Fonts catalogue, verified metric-compatible clones, weight-variant names, the CSS2 request and
 * the faces a document draws. DOM-free; `ooxml-ui`'s `fonts` folder does the browser side.
 */
export { GOOGLE_FONTS_CATALOGUE_DATE, GOOGLE_FONTS_FAMILIES } from './google-fonts-catalogue';
export { findGoogleFontsFamily } from './google-fonts-lookup';
export { findMetricCompatibleGoogleFontsFamily } from './metric-clones';
export * from './font-weight-names';
export * from './google-fonts-request';
export * from './webfont-faces';
