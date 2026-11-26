import { lookupPresetShape } from '../../geometry/preset-shape-evaluator';
import { anchorToPixelBox, pixelSizeToExtent } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import type { DrawingAnchor, DrawingObject, UnsupportedDrawingObject } from '../model';
import type { ShapeNode } from '../model-shapes';
import { type EditContext, sheetAt, sheetDrawingAt } from './context';
import { patchShapeText } from '../write/native-shape';
import { checkAnchor } from './drawings';
import { buildXml } from '../../xml/index';
import {
	parseShapeEquation,
	readEquationParagraphs,
	replaceShapeEquation,
	type ShapeEquation,
} from '../shape-equations';

export type NewDrawingShape =
	| { kind: 'preset'; preset: string; name: string }
	| { kind: 'textBox'; text: string; name: string }
	| { kind: 'equation'; equation: Required<ShapeEquation>; name: string }
	/** Ink to Shape: an unfilled preset outlined like the ink it replaces. */
	| {
			kind: 'outline';
			preset: string;
			name: string;
			line: import('../../drawingml/types').DrawingLine;
			rotation?: number;
			flipH?: boolean;
			flipV?: boolean;
	  };

export function addShape(
	ctx: EditContext,
	s: number,
	anchor: DrawingAnchor,
	value: NewDrawingShape,
): number {
	checkAnchor(anchor);
	const sheet = sheetAt(ctx.workbook, s);
	const textBox = value.kind === 'textBox' || value.kind === 'equation';
	if (value.kind === 'equation') parseShapeEquation(value.equation);
	const geometry = value.kind === 'preset' || value.kind === 'outline' ? value.preset : 'rect';
	if (!lookupPresetShape(geometry)) throw new RangeError('Unknown shape preset.');
	const isLine = geometry === 'line' || geometry === 'lineInv' || /Connector/.test(geometry);
	const extent =
		anchor.ext ??
		(() => {
			const box = anchorToPixelBox(sheet, anchor, createGridMetrics(sheet, { zoom: 100 }));
			return pixelSizeToExtent(box.w, box.h);
		})();
	const ink = { kind: 'scheme' as const, value: 'accent1', transforms: [] };
	const shape: ShapeNode = {
		kind: 'shape',
		name: value.name,
		geometry,
		frame: { x: 0, y: 0, width: extent.cx, height: extent.cy },
		fill:
			textBox || isLine || value.kind === 'outline'
				? { kind: 'none' }
				: { kind: 'solid', color: ink },
		...(value.kind === 'outline'
			? {
					line: structuredClone(value.line),
					...(value.rotation ? { rotation: value.rotation } : {}),
					...(value.flipH ? { flipH: true } : {}),
					...(value.flipV ? { flipV: true } : {}),
				}
			: {}),
		...(value.kind === 'outline'
			? {}
			: value.kind === 'equation'
				? {
						textBox: true,
						text: { text: '', paragraphs: [], insetsEmu: { left: 0, top: 0, right: 0, bottom: 0 } },
						equationParagraphs: [[{ kind: 'equation', ...value.equation }]],
					}
				: value.kind === 'textBox'
					? { textBox: true, text: plainText(value.text) }
					: { line: { fill: { kind: 'solid', color: ink }, widthEmu: 12700 } }),
	};
	return editDrawings(
		ctx,
		s,
		value.kind === 'equation'
			? 'Insert Equation'
			: value.kind === 'outline'
				? 'Ink to Shape'
				: textBox
					? 'Insert text box'
					: 'Insert shape',
		() =>
			sheet.drawings.push({
				kind: 'unsupported',
				description: textBox ? 'text box' : 'shape',
				anchor: structuredClone(anchor),
				shape,
			}) - 1,
	);
}

function plainText(text: string): NonNullable<ShapeNode['text']> {
	return { text, paragraphs: text.split(/\r?\n/).map((line) => ({ runs: [{ text: line }] })) };
}

export function editDrawings<T>(
	ctx: EditContext,
	sheet: number,
	label: string,
	change: () => T,
): T {
	return ctx.run(label, 'annotations', [{ kind: 'parts', sheet, parts: ['drawings'] }], change, {
		sheet,
	});
}

export function isWorksheetShape(
	drawing: DrawingObject | undefined,
): drawing is UnsupportedDrawingObject & { shape: ShapeNode } {
	return drawing?.kind === 'unsupported' && drawing.shape !== undefined;
}

export function setTextBoxText(ctx: EditContext, s: number, index: number, text: string): void {
	const drawing = sheetDrawingAt(ctx.workbook, s, index);
	if (
		!isWorksheetShape(drawing) ||
		!drawing.shape.textBox ||
		drawing.shape.wordArt ||
		drawing.shape.equationParagraphs
	)
		return;
	if ((drawing.shape.text?.text ?? '') === text) return;
	editDrawings(ctx, s, 'Edit text box', () => {
		if (drawing.shape.sourceShapeXml)
			drawing.shape.sourceShapeXml = patchShapeText(drawing.shape.sourceShapeXml, text);
		drawing.shape.text = plainText(text);
		if (drawing.sourceXml) drawing.shapeEdit = { ...drawing.shapeEdit, text: true };
	});
}

export function setShapeEquation(
	ctx: EditContext,
	sheet: number,
	index: number,
	equation: number,
	value: Required<ShapeEquation>,
): boolean {
	const drawing = sheetDrawingAt(ctx.workbook, sheet, index);
	if (!isWorksheetShape(drawing) || !drawing.shape.equationParagraphs) return false;
	const current = drawing.shape.equationParagraphs
		.flat()
		.filter((item) => item.kind === 'equation')[equation];
	if (!current?.omml) return false;
	if (current.omml === value.omml) return true;
	parseShapeEquation(value);
	if (drawing.sourceXml) {
		const changed = replaceShapeEquation(drawing.sourceXml, equation, value);
		if (!changed) return false;
		const next = readEquationParagraphs(changed.choice, changed.fallback);
		if (!next) return false;
		editDrawings(ctx, sheet, 'Edit Equation', () => {
			drawing.sourceXml = buildXml(changed.doc);
			drawing.shape.equationParagraphs = next;
		});
	} else {
		editDrawings(ctx, sheet, 'Edit Equation', () => Object.assign(current, value));
	}
	return true;
}
