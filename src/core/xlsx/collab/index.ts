// Real-time collaboration for workbooks (`ooxml-core/xlsx/collab`): the Yjs mapping of the
// spreadsheet model, a whole-model `DocumentAdapter` and the presence payload.
export {
	isSharedWorkbookEmpty,
	readSharedWorkbook,
	writeSharedWorkbook,
	xlsxDocumentAdapter,
	type XlsxAdapterOptions,
} from './adapter';
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
