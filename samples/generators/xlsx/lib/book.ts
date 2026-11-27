/**
 * A thin authoring layer over the core workbook model: a workbook with the sample fonts and a
 * custom theme palette, and cell formats described by a flat `Look` that is interned once.
 */
import type {
	Border,
	BorderEdge,
	BorderStyle,
	CellStyle,
	Color,
	Font,
	Workbook,
} from '../../../../src/core/xlsx/model';
import { internStyle } from '../../../../src/core/xlsx/styles';
import { createWorkbook } from '../../../../src/core/xlsx/workbook';
import { FONTS } from '../../lib/fiction';
import { Sheet } from './sheet';

/** One border edge: a colour (thin line) or a style with its colour. */
export type Edge = string | [BorderStyle, string];

/** A cell format, flat. Colours are `RRGGBB`. */
export interface Look {
	/** `display` is Rubik (the theme's major font); the default is Noto Sans. */
	font?: 'display' | 'body';
	size?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	color?: string;
	fill?: string;
	fmt?: string;
	h?: 'left' | 'center' | 'right' | 'centerContinuous';
	v?: 'top' | 'center' | 'bottom';
	wrap?: boolean;
	indent?: number;
	top?: Edge;
	bottom?: Edge;
	left?: Edge;
	right?: Edge;
	/** The cell stays editable on a protected sheet. */
	unlocked?: boolean;
	/** The name of the cell style the format derives from (see `Book.cellStyle`). */
	named?: string;
}

/** The twelve theme colours in SpreadsheetML order, as `RRGGBB`. */
export interface ThemeColors {
	light1: string;
	dark1: string;
	light2: string;
	dark2: string;
	accents: [string, string, string, string, string, string];
	link: string;
	followedLink: string;
}

export const rgb = (hex: string): Color => ({ rgb: `FF${hex.toUpperCase()}` });

const edge = (value: Edge): BorderEdge =>
	typeof value === 'string'
		? { style: 'thin', color: rgb(value) }
		: { style: value[0], color: rgb(value[1]) };

export class Book {
	readonly wb: Workbook;
	private readonly ids = new Map<string, number>();
	private readonly named = new Map<string, Look>();
	readonly sheets: Sheet[];

	constructor(
		names: string[],
		theme: ThemeColors,
		/** The colour of ordinary cell text. */
		readonly ink: string,
		readonly baseSize = 10,
	) {
		this.wb = createWorkbook({ sheets: names });
		this.wb.theme = {
			colors: [
				theme.light1,
				theme.dark1,
				theme.light2,
				theme.dark2,
				...theme.accents,
				theme.link,
				theme.followedLink,
			],
			majorFont: FONTS.display,
			minorFont: FONTS.body,
		};
		// Style 0 is what every unformatted cell and the row and column headers use.
		const normal = this.resolve({});
		this.wb.styles[0] = normal;
		this.wb.namedStyles = [{ name: 'Normal', style: normal, builtinId: 0 }];
		this.sheets = this.wb.sheets.map((ws, index) => new Sheet(this, ws, index));
	}

	sheet(name: string): Sheet {
		const found = this.sheets.find((sheet) => sheet.ws.name === name);
		if (!found) throw new Error(`No sheet named ${name}`);
		return found;
	}

	/** Registers a cell style of the gallery; `builtinId` ties it to Excel's built-in of that name. */
	cellStyle(name: string, look: Look, builtinId?: number): void {
		this.named.set(name, look);
		const entry = { name, style: this.resolve(look) };
		this.wb.namedStyles.push(builtinId === undefined ? entry : { ...entry, builtinId });
	}

	private font(look: Look): Font {
		const display = look.font === 'display';
		const font: Font = {
			name: display ? FONTS.display : FONTS.body,
			size: look.size ?? this.baseSize,
			family: 2,
			scheme: display ? 'major' : 'minor',
			color: rgb(look.color ?? this.ink),
		};
		if (look.bold) font.bold = true;
		if (look.italic) font.italic = true;
		if (look.underline) font.underline = 'single';
		return font;
	}

	private resolve(look: Look): CellStyle {
		const border: Border = {};
		if (look.top) border.top = edge(look.top);
		if (look.bottom) border.bottom = edge(look.bottom);
		if (look.left) border.left = edge(look.left);
		if (look.right) border.right = edge(look.right);
		const style: CellStyle = {
			font: this.font(look),
			fill: look.fill
				? { type: 'pattern', pattern: 'solid', fgColor: rgb(look.fill) }
				: { type: 'pattern', pattern: 'none' },
			border,
			numFmt: look.fmt ?? 'General',
		};
		if (look.h || look.v || look.wrap || look.indent)
			style.alignment = {
				...(look.h ? { horizontal: look.h } : {}),
				...(look.v ? { vertical: look.v } : {}),
				...(look.wrap ? { wrapText: true } : {}),
				...(look.indent ? { indent: look.indent } : {}),
			};
		if (look.unlocked) style.protection = { locked: false };
		if (look.named) style.cellStyleName = look.named;
		return style;
	}

	/** The style id of a look; a named look starts from its cell style. */
	style(look: Look | undefined): number {
		if (!look) return 0;
		const full = look.named ? { ...this.named.get(look.named), ...look } : look;
		const key = JSON.stringify(full, Object.keys(full).sort());
		let id = this.ids.get(key);
		if (id === undefined) {
			id = internStyle(this.wb, this.resolve(full));
			this.ids.set(key, id);
		}
		return id;
	}
}
