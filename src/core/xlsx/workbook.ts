import { OFFICE_COLORS } from '../drawingml/office-theme-colors';
import {
	OFFICE_THEME_FONTS,
	themeFontPair,
	type ThemeFontPair,
} from '../drawingml/office-theme-fonts';
import type { CellStyle, ThemePalette, Workbook, Worksheet } from './model';

/** The Office 2013+ default theme ("Office") palette, used when a package has no theme part. */
export const DEFAULT_THEME: ThemePalette = {
	colors: [...OFFICE_COLORS],
	majorFont: OFFICE_THEME_FONTS.major,
	minorFont: OFFICE_THEME_FONTS.minor,
};

/** Excel's default row height for 11pt Calibri, in points. */
export const DEFAULT_ROW_HEIGHT = 15;
/** Excel's default column width in characters (8.43 plus the 5px padding rounds to 8.43). */
export const DEFAULT_COL_WIDTH = 8.43;

/**
 * The default cell format of a new workbook (11pt in the theme's body font, General, no fill or
 * border). Excel names the face and marks it `scheme="minor"`, so it follows the theme.
 */
export function defaultCellStyle(minorFont: string = OFFICE_THEME_FONTS.minor): CellStyle {
	return {
		font: { name: minorFont, size: 11, family: 2, scheme: 'minor', color: { theme: 1 } },
		fill: { type: 'pattern', pattern: 'none' },
		border: {},
		numFmt: 'General',
	};
}

/** An empty worksheet with Excel's defaults. */
export function createWorksheet(name: string, sheetId: number): Worksheet {
	return {
		name,
		sheetId,
		state: 'visible',
		rows: new Map(),
		rowInfo: new Map(),
		columns: [],
		defaultRowHeight: DEFAULT_ROW_HEIGHT,
		merges: [],
		view: {
			showGridLines: true,
			showHeaders: true,
			showZeros: true,
			rightToLeft: false,
			zoom: 100,
		},
		hyperlinks: [],
		comments: [],
		conditionalFormats: [],
		dataValidations: [],
		tables: [],
		drawings: [],
		preserved: new Map(),
	};
}

export interface CreateWorkbookOptions {
	/** Sheet names to create; defaults to a single `Sheet1`. */
	sheets?: string[];
	/** The theme's heading and body fonts; by default `OFFICE_THEME_FONTS` (drawingml). */
	themeFonts?: Partial<ThemeFontPair>;
}

/** A new, empty workbook. */
export function createWorkbook(options: CreateWorkbookOptions = {}): Workbook {
	const names = options.sheets?.length ? options.sheets : ['Sheet1'];
	const fonts = themeFontPair(options.themeFonts);
	return {
		sheets: names.map((name, index) => createWorksheet(name, index + 1)),
		styles: [defaultCellStyle(fonts.minor)],
		namedStyles: [{ name: 'Normal', style: defaultCellStyle(fonts.minor), builtinId: 0 }],
		definedNames: [],
		theme: {
			...DEFAULT_THEME,
			colors: [...DEFAULT_THEME.colors],
			majorFont: fonts.major,
			minorFont: fonts.minor,
		},
		activeSheet: 0,
		date1904: false,
		properties: {},
		format: 'new',
		warnings: [],
	};
}

/** The sheet with `name` (case-insensitive, as Excel compares sheet names). */
export const sheetByName = (workbook: Workbook, name: string): Worksheet | undefined =>
	workbook.sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase());

/** Characters Excel forbids in sheet names. */
const INVALID_SHEET_CHARS = /[\\/?*[\]:]/;

/** Why a sheet name is invalid, or `undefined` when it is acceptable. */
export function validateSheetName(
	workbook: Workbook,
	name: string,
	exceptIndex = -1,
): string | undefined {
	if (!name.trim()) return 'A sheet name cannot be blank.';
	if (name.length > 31) return 'A sheet name can have at most 31 characters.';
	if (INVALID_SHEET_CHARS.test(name)) return 'A sheet name cannot contain \\ / ? * [ ] or :';
	if (name.startsWith("'") || name.endsWith("'"))
		return 'A sheet name cannot start or end with an apostrophe.';
	if (name.toLowerCase() === 'history') return '"History" is reserved by Excel.';
	const clash = workbook.sheets.findIndex(
		(sheet, index) => index !== exceptIndex && sheet.name.toLowerCase() === name.toLowerCase(),
	);
	return clash >= 0 ? 'That name is already taken.' : undefined;
}

/** The next free `SheetN` name. */
export function nextSheetName(workbook: Workbook, base = 'Sheet'): string {
	for (let n = workbook.sheets.length + 1; ; n++) {
		const name = `${base}${n}`;
		if (!sheetByName(workbook, name)) return name;
	}
}

/** The next unused `sheetId`. */
export const nextSheetId = (workbook: Workbook): number =>
	workbook.sheets.reduce((max, sheet) => Math.max(max, sheet.sheetId), 0) + 1;
