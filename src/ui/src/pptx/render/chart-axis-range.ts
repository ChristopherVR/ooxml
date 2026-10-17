// Compatibility facade: all products use the strict, format-neutral axis algorithms.
export {
	computeLogValueRange,
	computeValueRangeForAxis,
	computeStackedValueRangeForAxis,
	valueToYLog,
	findLogAxis,
	computeValueRangeForChart,
} from 'ooxml-core/chart';
export { generateLogTicks, generateAxisTicks, generateMinorAxisTicks } from 'ooxml-core/chart';
