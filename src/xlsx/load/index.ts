// Spreadsheet loading for every supported input: .xlsx/.xlsm (OPC zip), password-protected
// (encrypted) packages through the shared crypto area, Excel 97-2003 .xls and CSV. The legacy
// reader and the encrypted-package container inline the shared `@christophervr/ole2` codecs at
// build time (see tsup.config.ts); the main `xlsx` entry never imports ole2.
export {
	decodeText,
	detectWorkbookFormat,
	loadWorkbook,
	saveWorkbook,
	type LoadWorkbookOptions,
	type SaveWorkbookOptions,
	type WorkbookFormat,
} from './detect.js';
export {
	DataIntegrityError,
	IncorrectPasswordError,
	PasswordRequiredError,
	UnsupportedWorkbookError,
	isOoxmlCryptoError,
	type OoxmlCryptoErrorCode,
	type UnsupportedWorkbookErrorCode,
} from './errors.js';
export { isEncryptedOoxmlPackage, type EncryptionOptions } from '../../crypto/index.js';
export { LegacyXlsError, loadLegacyXls, type LegacyXlsErrorCode } from './legacy-xls.js';
export {
	csvToWorkbook,
	detectDelimiter,
	parseCsv,
	sheetNameFromFileName,
	neutraliseCsvFormula,
	sheetToCsv,
	type CsvOptions,
	type CsvWorkbookOptions,
	type CsvWriteOptions,
} from './csv.js';
export { decodeSpreadsheetText } from './text-decode.js';
