/** Browser-neutral text primitives shared by Office consumers. */
export {
	detectFontScript,
	segmentByScript,
	resolveFontForScript,
	hasDistinctScriptFonts,
	type FontScriptCategory,
	type ScriptRun,
} from './unicode-script-detection.js';
