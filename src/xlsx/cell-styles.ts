// Excel's built-in cell styles (Home > Cell Styles), as Excel 2013 and later define them for the
// Office theme. Each style lists only the format aspects it includes; applying it replaces those
// aspects and keeps the rest of the cell's format.
import type { Alignment, Border, CellStyle, Color, Fill, Font, Protection } from './model.js';

export interface BuiltinCellStyle {
	name: string;
	/** `builtinId` of `cellStyle` in styles.xml. */
	builtinId: number;
	/** Gallery group, as Excel shows them. */
	group: 'good-bad' | 'data' | 'titles' | 'themed' | 'number';
	/** Font changes over the workbook's Normal font (the whole font is replaced when present). */
	font?: Partial<Font>;
	fill?: Fill;
	border?: Border;
	numFmt?: string;
	alignment?: Alignment;
	protection?: Protection;
}

const rgb = (value: string): Color => ({ rgb: value });
const theme = (slot: number, tint?: number): Color =>
	tint ? { theme: slot, tint } : { theme: slot };
const solid = (fg: Color): Fill => ({
	type: 'pattern',
	pattern: 'solid',
	fgColor: fg,
	bgColor: { indexed: 64 },
});
const box = (style: 'thin' | 'double', color: Color): Border => ({
	left: { style, color },
	right: { style, color },
	top: { style, color },
	bottom: { style, color },
});

const ACCENTS: BuiltinCellStyle[] = [1, 2, 3, 4, 5, 6].flatMap((n): BuiltinCellStyle[] => {
	const slot = 3 + n;
	const base = 29 + (n - 1) * 4;
	return [
		{
			name: `Accent${n}`,
			builtinId: base,
			group: 'themed',
			font: { color: theme(0) },
			fill: solid(theme(slot)),
		},
		{
			name: `20% - Accent${n}`,
			builtinId: base + 1,
			group: 'themed',
			font: { color: theme(1) },
			fill: solid(theme(slot, 0.7999816888943144)),
		},
		{
			name: `40% - Accent${n}`,
			builtinId: base + 2,
			group: 'themed',
			font: { color: theme(1) },
			fill: solid(theme(slot, 0.5999938962981048)),
		},
		{
			name: `60% - Accent${n}`,
			builtinId: base + 3,
			group: 'themed',
			font: { color: theme(0) },
			fill: solid(theme(slot, 0.3999755851924192)),
		},
	];
});

/** Every built-in cell style this package knows, in gallery order. */
export const BUILTIN_CELL_STYLES: readonly BuiltinCellStyle[] = [
	{ name: 'Normal', builtinId: 0, group: 'good-bad' },
	{
		name: 'Bad',
		builtinId: 27,
		group: 'good-bad',
		font: { color: rgb('FF9C0006') },
		fill: solid(rgb('FFFFC7CE')),
	},
	{
		name: 'Good',
		builtinId: 26,
		group: 'good-bad',
		font: { color: rgb('FF006100') },
		fill: solid(rgb('FFC6EFCE')),
	},
	{
		name: 'Neutral',
		builtinId: 28,
		group: 'good-bad',
		font: { color: rgb('FF9C5700') },
		fill: solid(rgb('FFFFEB9C')),
	},
	{
		name: 'Calculation',
		builtinId: 22,
		group: 'data',
		font: { bold: true, color: rgb('FFFA7D00') },
		fill: solid(rgb('FFF2F2F2')),
		border: box('thin', rgb('FF7F7F7F')),
	},
	{
		name: 'Check Cell',
		builtinId: 23,
		group: 'data',
		font: { bold: true, color: theme(0) },
		fill: solid(rgb('FFA5A5A5')),
		border: box('double', rgb('FF3F3F3F')),
	},
	{
		name: 'Explanatory Text',
		builtinId: 53,
		group: 'data',
		font: { italic: true, color: rgb('FF7F7F7F') },
	},
	{
		name: 'Input',
		builtinId: 20,
		group: 'data',
		font: { color: rgb('FF3F3F76') },
		fill: solid(rgb('FFFFCC99')),
		border: box('thin', rgb('FF7F7F7F')),
	},
	{
		name: 'Linked Cell',
		builtinId: 24,
		group: 'data',
		font: { color: rgb('FFFA7D00') },
		border: { bottom: { style: 'double', color: rgb('FFFF8001') } },
	},
	{
		name: 'Note',
		builtinId: 10,
		group: 'data',
		fill: solid(rgb('FFFFFFCC')),
		border: box('thin', rgb('FFB2B2B2')),
	},
	{
		name: 'Output',
		builtinId: 21,
		group: 'data',
		font: { bold: true, color: rgb('FF3F3F3F') },
		fill: solid(rgb('FFF2F2F2')),
		border: box('thin', rgb('FF3F3F3F')),
	},
	{ name: 'Warning Text', builtinId: 11, group: 'data', font: { color: rgb('FFFF0000') } },
	{
		name: 'Heading 1',
		builtinId: 16,
		group: 'titles',
		font: { bold: true, size: 15, color: theme(3) },
		border: { bottom: { style: 'thick', color: theme(4) } },
	},
	{
		name: 'Heading 2',
		builtinId: 17,
		group: 'titles',
		font: { bold: true, size: 13, color: theme(3) },
		border: { bottom: { style: 'thick', color: theme(4, 0.499984740745262) } },
	},
	{
		name: 'Heading 3',
		builtinId: 18,
		group: 'titles',
		font: { bold: true, color: theme(3) },
		border: { bottom: { style: 'medium', color: theme(4, 0.3999755851924192) } },
	},
	{ name: 'Heading 4', builtinId: 19, group: 'titles', font: { bold: true, color: theme(3) } },
	{
		name: 'Title',
		builtinId: 15,
		group: 'titles',
		font: { size: 18, color: theme(3), scheme: 'major' },
	},
	{
		name: 'Total',
		builtinId: 25,
		group: 'titles',
		font: { bold: true, color: theme(1) },
		border: {
			top: { style: 'thin', color: theme(4) },
			bottom: { style: 'double', color: theme(4) },
		},
	},
	...ACCENTS,
	{
		name: 'Comma',
		builtinId: 3,
		group: 'number',
		numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
	},
	{
		name: 'Comma [0]',
		builtinId: 6,
		group: 'number',
		numFmt: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
	},
	{
		name: 'Currency',
		builtinId: 4,
		group: 'number',
		numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
	},
	{
		name: 'Currency [0]',
		builtinId: 7,
		group: 'number',
		numFmt: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
	},
	{ name: 'Percent', builtinId: 5, group: 'number', numFmt: '0%' },
	{
		name: 'Hyperlink',
		builtinId: 8,
		group: 'data',
		font: { color: theme(10), underline: 'single' },
	},
	{
		name: 'Followed Hyperlink',
		builtinId: 9,
		group: 'data',
		font: { color: theme(11), underline: 'single' },
	},
];

/** A built-in cell style by name (case-insensitive). */
export const builtinCellStyle = (name: string): BuiltinCellStyle | undefined =>
	BUILTIN_CELL_STYLES.find((s) => s.name.toLowerCase() === name.toLowerCase());

/**
 * Applies a built-in style's aspects to `base`; `normal` is the workbook's Normal format, whose
 * font the style's font changes start from. `Normal` itself resets every aspect.
 */
export function applyBuiltinStyle(
	base: CellStyle,
	entry: BuiltinCellStyle,
	normal: CellStyle,
): CellStyle {
	if (entry.builtinId === 0) {
		const reset = structuredClone(normal);
		delete reset.cellStyleName;
		return { ...reset, cellStyleName: 'Normal' };
	}
	const next: CellStyle = { ...structuredClone(base), cellStyleName: entry.name };
	if (entry.font) {
		const font: Font = { ...structuredClone(normal.font), ...structuredClone(entry.font) };
		next.font = font;
	}
	if (entry.fill) next.fill = structuredClone(entry.fill);
	if (entry.border) next.border = structuredClone(entry.border);
	if (entry.numFmt !== undefined) next.numFmt = entry.numFmt;
	if (entry.alignment) next.alignment = structuredClone(entry.alignment);
	if (entry.protection) next.protection = structuredClone(entry.protection);
	return next;
}
