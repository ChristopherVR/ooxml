import { fail } from './package-common';

interface Target {
	pageId: string;
	shapeId: string;
}
/** Whole-shape formatting. Font size and line weight use points. */
export interface VisioTextFormatEdit extends Target {
	type: 'format-text';
	fontSize?: number;
	fontFamily?: string;
	/** Opaque custom text color. */
	fontColor?: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strikethrough?: boolean;
	/** Enable paragraph bullets, retaining an existing glyph selection. */
	bullets?: boolean;
	/** Physical points, independent of the drawing scale. */
	indentLeft?: number;
	horizontalAlign?: 'left' | 'center' | 'right' | 'justify';
	verticalAlign?: 'top' | 'middle' | 'bottom';
}
export interface VisioShapeFormatEdit extends Target {
	type: 'format-shape';
	fillColor?: string;
	lineColor?: string;
	lineWeight?: number;
}
export type VisioFormatEdit = VisioTextFormatEdit | VisioShapeFormatEdit;
export const isVisioFormatEdit = (edit: { type: string }): edit is VisioFormatEdit =>
	edit.type === 'format-text' || edit.type === 'format-shape';

export function snapshotFormatting(edit: VisioFormatEdit): VisioFormatEdit {
	const result: VisioFormatEdit = { type: edit.type, pageId: edit.pageId, shapeId: edit.shapeId };
	const points = (value: unknown, minimum: number, maximum: number) => {
		if (typeof value !== 'number' || !Number.isFinite(value) || value < minimum || value > maximum)
			fail('INVALID_EDIT', 'Formatting point size is outside supported limits.');
		return value;
	};
	const color = (value: unknown, allowNone = false) => {
		if (
			typeof value !== 'string' ||
			(!(allowNone && value === 'none') && !/^#[0-9a-f]{6}$/i.test(value))
		)
			fail('INVALID_EDIT', 'Formatting colors require six hexadecimal digits.');
		return value.toLowerCase();
	};
	if (edit.type === 'format-text' && result.type === 'format-text') {
		if (edit.fontSize !== undefined) result.fontSize = points(edit.fontSize, 1, 1000);
		if (edit.fontColor !== undefined) result.fontColor = color(edit.fontColor);
		if (edit.indentLeft !== undefined) result.indentLeft = points(edit.indentLeft, 0, 7200);
		if (edit.fontFamily !== undefined) {
			if (
				typeof edit.fontFamily !== 'string' ||
				!edit.fontFamily.trim() ||
				edit.fontFamily.length > 256 ||
				/[\u0000-\u001f\u007f]/.test(edit.fontFamily)
			)
				fail('INVALID_EDIT', 'Invalid font family.');
			result.fontFamily = edit.fontFamily;
		}
		for (const name of ['bold', 'italic', 'underline', 'strikethrough', 'bullets'] as const) {
			if (edit[name] === undefined) continue;
			if (typeof edit[name] !== 'boolean') fail('INVALID_EDIT', 'Text styles require booleans.');
			result[name] = edit[name];
		}
		if (edit.horizontalAlign !== undefined) {
			if (!['left', 'center', 'right', 'justify'].includes(edit.horizontalAlign))
				fail('INVALID_EDIT', 'Invalid horizontal alignment.');
			result.horizontalAlign = edit.horizontalAlign;
		}
		if (edit.verticalAlign !== undefined) {
			if (!['top', 'middle', 'bottom'].includes(edit.verticalAlign))
				fail('INVALID_EDIT', 'Invalid vertical alignment.');
			result.verticalAlign = edit.verticalAlign;
		}
	} else if (edit.type === 'format-shape' && result.type === 'format-shape') {
		if (edit.fillColor !== undefined) result.fillColor = color(edit.fillColor, true);
		if (edit.lineColor !== undefined) result.lineColor = color(edit.lineColor);
		if (edit.lineWeight !== undefined) result.lineWeight = points(edit.lineWeight, 0, 100);
	}
	if (Object.keys(result).length === 3)
		fail('INVALID_EDIT', 'Formatting requires at least one property.');
	return result;
}
