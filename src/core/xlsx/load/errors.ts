// Typed load errors of `ooxml-core/xlsx/load`. Each has a stable `code` the UI can switch on (the
// classes are inlined into each subpath bundle, so compare `code` rather than `instanceof` when
// the error may come from another entry point).
export {
	DataIntegrityError,
	IncorrectPasswordError,
	PasswordRequiredError,
	isOoxmlCryptoError,
	type OoxmlCryptoErrorCode,
} from '../../crypto/errors.js';

export type UnsupportedWorkbookErrorCode = 'xlsb-unsupported';

/** The file is a spreadsheet format this library recognises but does not open. */
export class UnsupportedWorkbookError extends Error {
	public readonly code: UnsupportedWorkbookErrorCode;

	public constructor(code: UnsupportedWorkbookErrorCode, message: string) {
		super(message);
		this.name = 'UnsupportedWorkbookError';
		this.code = code;
	}
}

/** The .xlsb error: Excel Binary Workbook parts are BIFF12 records, not SpreadsheetML. */
export const xlsbUnsupported = (): UnsupportedWorkbookError =>
	new UnsupportedWorkbookError(
		'xlsb-unsupported',
		'Excel Binary Workbook (.xlsb) files are not supported; save as .xlsx in Excel.',
	);
