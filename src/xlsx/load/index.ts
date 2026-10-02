// Spreadsheet loading for every supported input: .xlsx/.xlsm (OPC zip), Excel 97-2003 .xls and
// CSV. The legacy reader inlines the shared `@christophervr/ole2` codecs at build time (see
// tsup.config.ts); the main `xlsx` entry never imports ole2.
export {
	decodeText,
	detectWorkbookFormat,
	loadWorkbook,
	saveWorkbook,
	type LoadWorkbookOptions,
	type WorkbookFormat,
} from './detect.js';
export { LegacyXlsError, loadLegacyXls, type LegacyXlsErrorCode } from './legacy-xls.js';
export {
	csvToWorkbook,
	detectDelimiter,
	parseCsv,
	sheetNameFromFileName,
	sheetToCsv,
	type CsvOptions,
	type CsvWorkbookOptions,
} from './csv.js';
