// Real-time collaboration for workbooks (`ooxml-core/xlsx/collab`): the Yjs mapping of the
// spreadsheet model, a whole-model `DocumentAdapter`, the per-edit session binding and presence.
export {
	isSharedWorkbookEmpty,
	readSharedWorkbook,
	writeSharedWorkbook,
	xlsxDocumentAdapter,
	type XlsxAdapterOptions,
} from './adapter';
export {
	XLSX_EDIT_ORIGIN,
	XLSX_SYNC_ORIGIN,
	bindWorkbookSession,
	scopeOf,
	type WorkbookBinding,
	type WorkbookBindingOptions,
	type WorkbookCollabHost,
} from './bind';
export type { CellEntry, LineEntry } from './codec';
export type { NameEntry } from './names';
export {
	resolveSelections,
	sanitizeXlsxPresence,
	type RemoteSelection,
	type XlsxPresence,
} from './presence';
export { META, NAMES, SCHEMA_VERSION, SHEETS, STYLES, sharedWorkbook } from './schema';
export type { SharedWorkbook } from './schema';
export { sharedStyleKey } from './styles';
export type { WriteScope } from './write';
