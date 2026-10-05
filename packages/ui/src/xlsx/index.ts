// `<xlsx-editor>`: the Excel-style spreadsheet editor element (title bar, ribbon, formula bar,
// grid, sheet tabs, status bar, File backstage). The framework bindings, the demo and the
// published packages import from here; keep the names stable.
export { XlsxEditorElement, defineXlsxEditor, type XlsxThemeColors } from './component';
export {
	XLSX_EDITOR_EVENTS,
	type FileCommand,
	type FileCommandDetail,
	type SelectionChangeDetail,
	type XlsxEditorEventDetail,
	type XlsxEditorEventMap,
	type XlsxEditorEventName,
} from './events';
export {
	XLSX_EDITOR_ATTRIBUTES,
	DEFAULT_AUTHOR_NAME,
	DEFAULT_FILE_NAME,
} from './editor-attributes';
export {
	EDITOR_LOCALES,
	normalizeEditorLocale,
	type EditorLocale,
	type EditorLocaleInput,
} from './localization';
export {
	THEME_KEYS,
	darkTheme as xlsxDarkTheme,
	lightTheme as xlsxLightTheme,
	themeToCssVars,
	type EditorThemeMode,
	type XlsxTheme,
} from './theme';
// Extension points: commands, dialogs and ribbon tabs added by a host.
export type { Command as EditorCommand, CommandRegistry } from './commands';
export type {
	EditorContext,
	GridController,
	Selection as EditorSelection,
	SelectionModel,
} from './context';
export type { DialogRegistry } from './dialogs';
export {
	registerRibbonTabs,
	type RibbonControl,
	type RibbonGroup,
	type RibbonMenuItem,
	type RibbonTab,
} from './ribbon/parts';
export { registerRibbonIcon } from './ribbon/icons';
