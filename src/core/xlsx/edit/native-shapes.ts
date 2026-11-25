import { lookupPresetShape } from '../../geometry/preset-shape-evaluator';
import { anchorToPixelBox, pixelSizeToExtent } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import type { DrawingAnchor, DrawingObject, UnsupportedDrawingObject } from '../model';
import type { ShapeNode } from '../model-shapes';
import { type EditContext, sheetAt, sheetDrawingAt } from './context';
import { patchShapeText } from '../write/native-shape';
import { checkAnchor } from './drawings';

export type NewDrawingShape =
	| { kind: 'preset'; preset: string; name: string }
	| { kind: 'textBox'; text: string; name: string };

export function addShape(
	ctx: EditContext,
	s: number,
	anchor: DrawingAnchor,
	value: NewDrawingShape,
): number {
	checkAnchor(anchor);
	const sheet = sheetAt(ctx.workbook, s);
	const textBox = value.kind === 'textBox';
	const geometry = value.kind === 'preset' ? value.preset : 'rect';
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
		fill: textBox || isLine ? { kind: 'none' } : { kind: 'solid', color: ink },
		...(textBox
			? { textBox: true, text: plainText(value.text) }
			: { line: { fill: { kind: 'solid', color: ink }, widthEmu: 12700 } }),
	};
	return editDrawings(
		ctx,
		s,
		textBox ? 'Insert text box' : 'Insert shape',
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
	if (!isWorksheetShape(drawing) || !drawing.shape.textBox || drawing.shape.wordArt) return;
	if ((drawing.shape.text?.text ?? '') === text) return;
	editDrawings(ctx, s, 'Edit text box', () => {
		if (drawing.shape.sourceShapeXml)
			drawing.shape.sourceShapeXml = patchShapeText(drawing.shape.sourceShapeXml, text);
		drawing.shape.text = plainText(text);
		if (drawing.sourceXml) drawing.shapeEdit = { ...drawing.shapeEdit, text: true };
	});
}
