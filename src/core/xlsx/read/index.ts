export { loadXlsx, type LoadXlsxOptions } from './load';
export {
	DEFAULT_XLSX_LOAD_LIMITS,
	WorkbookTooLargeError,
	isWorkbookTooLargeError,
	type XlsxLoadLimits,
} from './limits';
export { translateReferences, stripFuturePrefixes, addFuturePrefixes } from './formula-text';
export { readSparklineGroup, readSparklineGroups, sheetSparklineGroups } from './sparklines';
