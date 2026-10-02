// Surface every published framework package re-exports next to its component, so an application
// needs no second package to build a workbook model, configure the editor or open an .xlsx,
// legacy .xls or .csv file (the model API comes from `@christophervr/xlsx-core`, which installs
// with the editor).
export * from '@christophervr/xlsx-core';
export * from 'xlsx-web-component';
export {
	csvToWorkbook,
	detectWorkbookFormat,
	LegacyXlsError,
	loadLegacyXls,
	loadWorkbook,
	parseCsv,
	saveWorkbook,
	sheetToCsv,
} from '@christophervr/xlsx-core/load';
export type { WorkbookFormat } from '@christophervr/xlsx-core/load';
export type {
	EditorBinding,
	EditorEventOptions,
	EditorHandle,
	EditorOptions,
	EditorProps,
} from './index';
