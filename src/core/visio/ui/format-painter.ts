import type { VisioDocument, VisioPage } from '../model';
import type { VisioShapeFormatEdit, VisioTextFormatEdit } from '../edit-formatting-commands';
import { visioFontFamilies, visioStyleFormattingShape } from './formatting';
import { visioShapeFormattingState } from './shape-formatting';

type Patch<T> = Omit<T, 'type' | 'pageId' | 'shapeId'>;
/** Formatting captured by Format Painter, expressed only as existing whole-shape format edits. */
export interface VisioFormatPainterSnapshot {
	readonly shape: Readonly<Patch<VisioShapeFormatEdit>>;
	readonly text: Readonly<Patch<VisioTextFormatEdit>>;
	/** Human-readable source formatting that existing edits cannot express and was not copied. */
	readonly skipped: readonly string[];
}
/** Never copied: no core edit expresses them yet. */
export const VISIO_FORMAT_PAINTER_LIMITS =
	'Shadows, rounding, theme effects, arrowheads and per-run text differences are not copied.';

const hex = (value: string | undefined): value is string =>
	typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value);
const round = (value: number) => Math.round(value * 100) / 100;

/**
 * Reads the source shape's effective fill, line and text formatting from the scene. Text uses
 * the first run and paragraph, as Visio's painter applies one character and paragraph format.
 */
export function visioFormatPainterSnapshot(
	document: VisioDocument,
	page: VisioPage,
	shapeId: string,
): VisioFormatPainterSnapshot | undefined {
	const source = visioStyleFormattingShape(page, shapeId);
	if (!source) return undefined;
	const style = source.style,
		paint = visioShapeFormattingState([source]),
		skipped: string[] = [];
	const shape: Patch<VisioShapeFormatEdit> = {};
	if (style.fillGradient) skipped.push('gradient fill');
	else if (style.fill === 'none' || paint.fillPatternIndex === 0) shape.fillColor = 'none';
	else {
		if (hex(style.fill)) shape.fillColor = style.fill;
		const pattern = paint.fillPatternIndex;
		if (pattern !== undefined && Number.isInteger(pattern) && pattern >= 1 && pattern <= 24)
			shape.fillPattern = pattern;
		else if (pattern !== undefined) skipped.push('fill pattern');
		if (hex(paint.fillBackgroundColor)) shape.fillBackgroundColor = paint.fillBackgroundColor;
		if (paint.fillTransparency !== undefined) shape.fillTransparency = paint.fillTransparency;
		else if (paint.fillForegroundTransparency !== undefined)
			skipped.push('separate fill transparencies');
	}
	if (style.lineGradient) skipped.push('gradient line');
	else if (hex(style.lineColor)) shape.lineColor = style.lineColor;
	const weight = round(style.lineWidth * 72);
	if (Number.isFinite(weight) && weight >= 0 && weight <= 100) shape.lineWeight = weight;
	else skipped.push('line weight');
	const pattern = style.linePattern;
	if (Number.isInteger(pattern) && pattern >= 0 && pattern <= 23) shape.linePattern = pattern;
	else skipped.push('line pattern');
	if (!style.lineGradient && paint.lineTransparency !== undefined)
		shape.lineTransparency = paint.lineTransparency;
	if (style.startArrow || style.endArrow) skipped.push('arrowheads');

	const text: Patch<VisioTextFormatEdit> = {};
	const runs = source.text.runs;
	const run = runs[0] ?? source.text;
	if (runs.length > 1) {
		const keys = ['fontFamily', 'fontSize', 'color', 'bold', 'italic', 'underline'] as const;
		if (runs.some((other) => keys.some((key) => other[key] !== run[key])))
			skipped.push('mixed text formatting (the first run is used)');
	}
	if (visioFontFamilies(document).includes(run.fontFamily)) text.fontFamily = run.fontFamily;
	else skipped.push('font');
	const size = round(run.fontSize * 72);
	if (size >= 1 && size <= 1000) text.fontSize = size;
	if (hex(run.color)) text.fontColor = run.color;
	if (run.opacity !== undefined && run.opacity < 1) skipped.push('text transparency');
	text.bold = !!run.bold;
	text.italic = !!run.italic;
	text.underline = !!run.underline;
	text.strikethrough = !!run.strikethrough;
	const paragraph = source.text.paragraphs?.[0];
	const align = paragraph?.horizontalAlign ?? source.text.horizontalAlign;
	if (align === 'distributed') skipped.push('distributed alignment');
	else text.horizontalAlign = align;
	text.verticalAlign = source.text.verticalAlign;
	if (paragraph) {
		const indent = round(paragraph.indentLeft * 72);
		if (indent >= 0 && indent <= 7200) text.indentLeft = indent;
		text.bullets = !!paragraph.bullet;
	}
	return Object.freeze({
		shape: Object.freeze(shape),
		text: Object.freeze(text),
		skipped: Object.freeze(skipped),
	});
}

/**
 * One format-shape and one format-text edit per target; undefined when a target is ineligible.
 * Text formatting needs existing local text (core refuses to create a Text element for it), so
 * targets without text receive only fill and line formatting.
 */
export function visioFormatPainterEdits(
	page: VisioPage,
	snapshot: VisioFormatPainterSnapshot,
	shapeIds: readonly string[],
): (VisioShapeFormatEdit | VisioTextFormatEdit)[] | undefined {
	const shapes = shapeIds.map((id) => visioStyleFormattingShape(page, id));
	if (!shapes.length || shapes.some((shape) => !shape)) return undefined;
	return shapes.flatMap((shape) => {
		const target = { pageId: page.id, shapeId: shape!.id };
		return [
			...(Object.keys(snapshot.shape).length
				? [{ type: 'format-shape' as const, ...target, ...snapshot.shape }]
				: []),
			...(Object.keys(snapshot.text).length && shape!.text.plainText
				? [{ type: 'format-text' as const, ...target, ...snapshot.text }]
				: []),
		];
	});
}
