/**
 * Editor theme tokens.
 *
 * Provenance: the semantic key set mirrors docx-viewer's theme (packages/web-component/src/theme,
 * itself following pptx-viewer's shared `ViewerThemeColors` and the shadcn/ui naming), with the
 * `--xve-*` prefix, plus grid tokens for the spreadsheet surface. Keys are camelCase here and
 * become kebab-case custom properties. Values accept any CSS colour. Cells stay white like Excel's
 * sheet in every theme; only the chrome (and the row and column headers) follow dark mode.
 */
export interface XlsxTheme {
	/** Application chrome background (title bar, status bar). */
	background: string;
	foreground: string;
	/** Ribbon, panel and backstage surface. */
	card: string;
	cardForeground: string;
	/** Menus, dialogs and other floating surfaces. */
	popover: string;
	popoverForeground: string;
	/** Brand and primary action colour (Excel green). */
	primary: string;
	primaryForeground: string;
	secondary: string;
	secondaryForeground: string;
	muted: string;
	mutedForeground: string;
	/** Hover and selected highlight. */
	accent: string;
	accentForeground: string;
	destructive: string;
	destructiveForeground: string;
	border: string;
	input: string;
	/** Focus ring. */
	ring: string;
	/** Base corner radius, e.g. `4px`. */
	radius: string;
	/** Cell gridlines. */
	gridLine: string;
	/** Row and column header background and text. */
	headerBg: string;
	headerFg: string;
	/** Header of a selected row or column. */
	headerActiveBg: string;
	/** Fill over selected cells. */
	selection: string;
	/** Outline of the selection and the active cell. */
	selectionBorder: string;
	/** The sheet behind the cells (white like Excel's paper). */
	sheetBg: string;
}

export type EditorThemeMode = 'light' | 'dark' | 'auto';
export const EDITOR_THEME_MODES: readonly EditorThemeMode[] = ['light', 'dark', 'auto'];
