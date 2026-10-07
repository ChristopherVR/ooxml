/** Compatibility entry: canonical equation logic lives in ooxml-core/math. */
export {
	type LatexParserContext,
	type ScriptArgs,
	parseScriptArgs,
	applyScripts,
	tryParseScripts,
	parseNary,
	parseDelimiter,
	parseFuncApplication,
	parseTextArgument,
} from 'ooxml-core/math';
