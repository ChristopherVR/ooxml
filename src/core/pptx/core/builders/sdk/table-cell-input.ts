/**
 * The cell style of a table cell made in code: its text style, its own fill
 * and its edge lines, mapped onto the {@link PptxTableCellStyle} fields the
 * save path writes into `a:tcPr`.
 *
 * @module sdk/table-cell-input
 */

import type { PptxTableCellStyle } from '../../types/table';
import { tableCellBorderStyle } from './table-border-input';
import type { FillInput, TableCellInput } from './types';

/** The parts of a gradient a cell's `a:gradFill` has no place for. */
function refuseGradientExtras(fill: Extract<FillInput, { type: 'gradient' }>): void {
	const radial = fill.path !== undefined || fill.gradientType === 'radial';
	const unsupported = [
		fill.tile !== undefined ? 'tile' : undefined,
		fill.scaled === false ? 'scaled: false' : undefined,
		fill.rotWithShape !== undefined ? 'rotWithShape' : undefined,
		fill.focus !== undefined && !radial ? 'focus on a linear gradient' : undefined,
	].filter((name): name is string => name !== undefined);
	if (unsupported.length > 0) {
		throw new RangeError(
			`TableCellInput.fill: ${unsupported.join(', ')} cannot be written on a table cell`,
		);
	}
	if (fill.stops.length === 0) {
		throw new RangeError('TableCellInput.fill: a gradient needs at least one stop');
	}
}

/**
 * A cell's own fill as cell style fields.
 *
 * @throws RangeError for a fill a cell cannot carry: an image, a solid fill
 *   with an opacity below 1, a pattern with no preset, a gradient with no
 *   stop, or a gradient's `tile`, `scaled: false`, `rotWithShape` or (on a
 *   linear gradient) `focus`.
 */
export function tableCellFillStyle(fill: FillInput | undefined): Partial<PptxTableCellStyle> {
	switch (fill?.type) {
		case undefined:
			return {};
		case 'solid':
			if (fill.opacity !== undefined && fill.opacity !== 1) {
				throw new RangeError(
					`TableCellInput.fill.opacity cannot be written on a table cell, got ${fill.opacity}`,
				);
			}
			return {
				fillMode: 'solid',
				backgroundColor: fill.color,
				...(fill.themeColorRef ? { backgroundColorRef: fill.themeColorRef } : {}),
			};
		case 'gradient': {
			refuseGradientExtras(fill);
			const radial = fill.path !== undefined || fill.gradientType === 'radial';
			return {
				fillMode: 'gradient',
				gradientFillType: radial ? 'radial' : 'linear',
				gradientFillAngle: fill.angle ?? 0,
				gradientFillStops: fill.stops.map((stop) => ({
					color: stop.color,
					position: stop.position,
					...(stop.opacity !== undefined ? { opacity: stop.opacity } : {}),
				})),
				...(radial ? { gradientFillPathType: fill.path ?? 'circle' } : {}),
				...(fill.focus ? { gradientFillFillToRect: { ...fill.focus } } : {}),
			};
		}
		case 'pattern':
			if (!fill.preset) {
				throw new RangeError('TableCellInput.fill: a pattern fill needs a preset');
			}
			return {
				fillMode: 'pattern',
				patternFillPreset: fill.preset,
				...(fill.foreground !== undefined ? { patternFillForeground: fill.foreground } : {}),
				...(fill.background !== undefined ? { patternFillBackground: fill.background } : {}),
			};
		case 'image':
			throw new RangeError('TableCellInput.fill: an image fill cannot be written on a table cell');
		case 'none':
			return { fillMode: 'none' };
	}
}

/** The cell's text style as cell style fields. */
function tableCellTextStyle(style: TableCellInput['style']): PptxTableCellStyle {
	const out: PptxTableCellStyle = {};
	if (style?.fontSize !== undefined) out.fontSize = style.fontSize;
	if (style?.bold !== undefined) out.bold = style.bold;
	if (style?.italic !== undefined) out.italic = style.italic;
	if (style?.color !== undefined) out.color = style.color;
	if (style?.alignment !== undefined) out.align = style.alignment;
	return out;
}

/** The cell style of a cell made in code, or `undefined` when it sets none. */
export function tableCellStyle(cell: TableCellInput): PptxTableCellStyle | undefined {
	const style: PptxTableCellStyle = {
		...tableCellTextStyle(cell.style),
		...tableCellFillStyle(cell.fill),
		...tableCellBorderStyle(cell.borders),
	};
	return cell.style || Object.keys(style).length > 0 ? style : undefined;
}
