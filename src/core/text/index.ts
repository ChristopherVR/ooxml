/** Browser-neutral text primitives shared by Office consumers. */
export {
	detectFontScript,
	segmentByScript,
	resolveFontForScript,
	hasDistinctScriptFonts,
	type FontScriptCategory,
	type ScriptRun,
} from './unicode-script-detection';
export * from './date-picture';
export * from './tab-leader';
export * from './decimal-tab';
export * from './wrap-styled-runs';
export * from './change-case';
export * from './merge-three-way';
export * from './font-metrics/advance-width';
export * from './font-metrics/vertical-metrics';
export { parseTsv, toTsv } from './tsv';
export * from './locale-number';
export * from './realign-runs';

export * from './font-substitution';
export * from './metric-font-family';
export * from './script-fonts';
export { parseTrueType, type TrueTypeFont } from './truetype/parse';
export { subsetTrueType, type TrueTypeSubset } from './truetype/subset';
export * from './webfonts';
