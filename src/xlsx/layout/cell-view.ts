import { rangeContains } from '../address.js';
import { getCell } from '../cells.js';
import type { Alignment, CellStyle, CellValue, Workbook, Worksheet } from '../model.js';
import { isCellError } from '../model.js';
import { formatValue } from '../numfmt/index.js';
import { styleAt } from '../styles.js';
import { bordersView, fillView, fontView, mergeBorder, mergeFont } from './style-view.js';
import type { CellView, ConditionalFormatEvaluator, HAlignView } from './types.js';
import { INDENT_PX_PER_LEVEL } from './units.js';

/** The style id that applies to a cell: its own, else its row's, else its column's. */
export function effectiveStyleId(sheet: Worksheet, row: number, col: number): number {
	const cell = getCell(sheet, row, col);
	if (cell?.styleId !== undefined) return cell.styleId;
	if (cell) return 0;
	const rowStyle = sheet.rowInfo.get(row)?.styleId;
	if (rowStyle !== undefined) return rowStyle;
	const column = sheet.columns.find((c) => col >= c.min && col <= c.max);
	return column?.styleId ?? 0;
}

/** Excel's `General` horizontal alignment: numbers right, booleans and errors centred, text left. */
export function generalAlignment(value: CellValue): 'left' | 'center' | 'right' {
	if (typeof value === 'number') return 'right';
	if (typeof value === 'boolean' || isCellError(value)) return 'center';
	return 'left';
}

/** Converts `textRotation` (0-90 up, 91-180 down, 255 stacked) to degrees counter-clockwise. */
export function rotationDegrees(textRotation: number | undefined): {
	rotation: number;
	vertical: boolean;
} {
	if (textRotation === undefined || textRotation === 0) return { rotation: 0, vertical: false };
	if (textRotation === 255) return { rotation: 0, vertical: true };
	if (textRotation > 90 && textRotation <= 180)
		return { rotation: -(textRotation - 90), vertical: false };
	return { rotation: Math.max(-90, Math.min(90, textRotation)), vertical: false };
}

export interface CellViewOptions {
	/** Pixels per indent level (default 9, Excel's step with 11pt Calibri). */
	indentPx?: number;
}

function displayText(
	value: CellValue,
	style: CellStyle,
	numFmt: string,
	workbook: Workbook,
	sheet: Worksheet,
): { text: string; color?: string } {
	if (value === null || value === '') return { text: '' };
	if (typeof value === 'number' && value === 0 && !sheet.view.showZeros) return { text: '' };
	if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };
	if (isCellError(value)) return { text: value.error };
	const formatted = formatValue(value, numFmt || style.numFmt || 'General', {
		date1904: workbook.date1904,
	});
	return formatted.color
		? { text: formatted.text, color: formatted.color }
		: { text: formatted.text };
}

/**
 * Everything the grid needs to paint one cell: display text, resolved font, fill, borders,
 * alignment and conditional-format decorations. DOM-free; text is never measured here.
 */
export function cellView(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
	cf?: ConditionalFormatEvaluator,
	options: CellViewOptions = {},
): CellView {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
	const cell = getCell(sheet, row, col);
	const value: CellValue = cell?.value ?? null;
	const style = styleAt(workbook, effectiveStyleId(sheet, row, col));
	const conditional = cf?.at(row, col);
	const dxf = conditional?.style;

	// A cell font is a complete spec; the default style only fills a missing name or size.
	const defaultFont = workbook.styles[0]?.font ?? {};
	const base = { ...style.font };
	if (!base.name && !base.scheme) {
		if (defaultFont.name) base.name = defaultFont.name;
		if (defaultFont.scheme) base.scheme = defaultFont.scheme;
	}
	if (!base.size && defaultFont.size) base.size = defaultFont.size;
	const font = mergeFont(base, dxf?.font);
	// Show Formulas mode (Ctrl+`): formula cells show their text, left aligned, unformatted.
	const formulaText =
		sheet.view.showFormulas && cell?.formula !== undefined ? `=${cell.formula}` : undefined;
	const shown =
		formulaText !== undefined
			? { text: formulaText }
			: displayText(value, style, dxf?.numFmt ?? style.numFmt, workbook, sheet);
	const fontViewResolved = fontView(font, workbook.theme);
	if (shown.color) fontViewResolved.color = shown.color;

	const alignment: Alignment = style.alignment ?? {};
	const horizontal = alignment.horizontal ?? 'general';
	const hAlign: HAlignView =
		horizontal !== 'general'
			? horizontal
			: formulaText !== undefined
				? 'left'
				: generalAlignment(value);
	const { rotation, vertical } = rotationDegrees(alignment.textRotation);
	const wrap = alignment.wrapText === true || hAlign === 'justify' || hAlign === 'distributed';
	const shrink = alignment.shrinkToFit === true && !wrap;
	const indentable = hAlign === 'left' || hAlign === 'right' || hAlign === 'distributed';

	let fill = dxf?.fill ? fillView(dxf.fill, workbook.theme, true) : undefined;
	if (!fill) fill = fillView(style.fill, workbook.theme);
	if (conditional?.colorScale && !dxf?.fill) fill = { background: conditional.colorScale };

	const isNumber = typeof value === 'number' && formulaText === undefined;
	const view: CellView = {
		text: conditional?.hideValue ? '' : shown.text,
		font: fontViewResolved,
		borders: bordersView(mergeBorder(style.border, dxf?.border), workbook.theme),
		hAlign,
		vAlign: alignment.vertical ?? 'bottom',
		wrap,
		shrink,
		indentPx: indentable ? (alignment.indent ?? 0) * (options.indentPx ?? INDENT_PX_PER_LEVEL) : 0,
		rotation,
		isNumber,
		overflow:
			(formulaText !== undefined || (typeof value === 'string' && value !== '')) &&
			!wrap &&
			!shrink &&
			rotation === 0 &&
			!vertical &&
			hAlign !== 'fill' &&
			!sheet.merges.some((m) => rangeContains(m, { row, col })),
		hasComment: sheet.comments.some((c) => c.address.row === row && c.address.col === col),
		hasHyperlink: false,
		validationList: sheet.dataValidations.some(
			(dv) =>
				dv.type === 'list' &&
				dv.showDropDown !== false &&
				dv.ranges.some((r) => rangeContains(r, { row, col })),
		),
		isError: isCellError(value) && formulaText === undefined,
	};
	if (vertical) view.verticalText = true;
	if (fill) view.fill = fill;
	if (conditional?.dataBar) view.dataBar = conditional.dataBar;
	if (conditional?.icon) view.icon = conditional.icon;

	const link = sheet.hyperlinks.find((h) => rangeContains(h.range, { row, col }));
	if (link) {
		view.hasHyperlink = true;
		const title = link.tooltip ?? link.target ?? link.location;
		if (title) view.title = title;
	}

	if (cell?.richText?.length && typeof value === 'string' && !conditional?.hideValue) {
		view.rich = cell.richText.map((run) => ({
			text: run.text,
			font: fontView(mergeFont(mergeFont(font, run.font), dxf?.font), workbook.theme),
		}));
	}
	return view;
}
