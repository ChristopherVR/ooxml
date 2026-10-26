// A generated 2D cache, independent of Office's native layout regeneration. This is not a
// round-trip writer for a source cached part: unsupported 3D must remain in its original part.
import { createOfficeGuid } from '../../crypto/uuid';
import { normalizeStShapeType } from '../../geometry/preset-geometry-names';
import { drawingFillXml } from '../../drawingml/write-fill';
import { drawingLineXml } from '../../drawingml/write-line';
import { drawingEffectsXml } from '../../drawingml/drawing-effects';
import { drawingRichTextBodyXml } from '../../drawingml/write-text-body';
import type { DrawingRichTextBody, DrawingRunProperties } from '../../drawingml/text-model';
import type { DiagramDrawing, DiagramDrawingShape, DiagramPath } from '../types';
import { DIAGRAM_XML_PROLOG, diagramAvLstXml, escapeDiagramXml as esc } from './data';

const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const GUID = /^\{[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}\}$/i;
function pathXml(path: DiagramPath): string {
	const point = (x: number, y: number) => `<a:pt x="${Math.round(x)}" y="${Math.round(y)}"/>`;
	const commands = path.commands
		.map((command) => {
			switch (command.op) {
				case 'M':
					return `<a:moveTo>${point(command.x, command.y)}</a:moveTo>`;
				case 'L':
					return `<a:lnTo>${point(command.x, command.y)}</a:lnTo>`;
				case 'C':
					return `<a:cubicBezTo>${point(command.x1, command.y1)}${point(command.x2, command.y2)}${point(command.x, command.y)}</a:cubicBezTo>`;
				case 'Q':
					return `<a:quadBezTo>${point(command.x1, command.y1)}${point(command.x, command.y)}</a:quadBezTo>`;
				case 'A':
					return `<a:arcTo wR="${Math.round(command.wR)}" hR="${Math.round(command.hR)}" stAng="${Math.round(command.startAngle * 60000)}" swAng="${Math.round(command.swingAngle * 60000)}"/>`;
				case 'Z':
					return '<a:close/>';
			}
		})
		.join('');
	return `<a:path${path.width !== undefined ? ` w="${Math.round(path.width)}"` : ''}${path.height !== undefined ? ` h="${Math.round(path.height)}"` : ''}${path.fill ? ` fill="${esc(path.fill)}"` : ''}${path.stroke === false ? ' stroke="0"' : ''}>${commands}</a:path>`;
}

function textBody(shape: DiagramDrawingShape): DrawingRichTextBody | undefined {
	if (shape.body) return shape.body;
	if (!shape.text) return undefined;
	return {
		bodyProperties: { ...(shape.text.anchor ? { anchor: shape.text.anchor } : {}) },
		paragraphs: shape.text.paragraphs.map((paragraph) => ({
			...(paragraph.align ? { properties: { align: paragraph.align } } : {}),
			items: paragraph.runs.map((run) => {
				const properties: DrawingRunProperties = {
					...(run.sizePt !== undefined ? { sizePt: run.sizePt } : {}),
					...(run.bold !== undefined ? { bold: run.bold } : {}),
					...(run.italic !== undefined ? { italic: run.italic } : {}),
					...(run.underline !== undefined ? { underline: run.underline ? 'sng' : 'none' } : {}),
					...(run.color ? { fill: { kind: 'solid', color: run.color } as const } : {}),
					...(run.typeface ? { latin: { typeface: run.typeface } } : {}),
				};
				return { kind: 'run' as const, text: run.text, properties };
			}),
		})),
	};
}

/** Serialize generated neutral shapes, returning the GUIDs a host can bind to presentation points. */
export function writeGeneratedDiagramDrawing(
	drawing: DiagramDrawing,
	options: {
		createModelId?: () => string;
		presentationIds?: ReadonlyMap<string, string>;
	} = {},
): { xml: string; modelIds: ReadonlyMap<string, string> } {
	const modelIds = new Map<string, string>();
	const body = drawing.shapes
		.map((shape, index) => {
			if (shape.has3d) throw new Error('A generated 2D drawing cannot serialize 3D SmartArt.');
			if (!Object.values(shape.frame).every(Number.isFinite))
				throw new Error('The generated diagram has invalid geometry.');
			const id =
				options.presentationIds?.get(shape.modelId) ??
				(GUID.test(shape.modelId) ? shape.modelId : (options.createModelId ?? createOfficeGuid)());
			if (!GUID.test(id)) throw new Error('A diagram drawing requires GUID presentation ids.');
			modelIds.set(shape.modelId, id);
			const text = textBody(shape);
			return `<dsp:sp modelId="${esc(id)}"><dsp:nvSpPr><dsp:cNvPr id="${index + 1}" name="Shape ${index + 1}"/><dsp:cNvSpPr/></dsp:nvSpPr>${diagramShapePropertiesXml(shape)}${text ? drawingRichTextBodyXml(text, 'dsp:txBody', DSP) : ''}</dsp:sp>`;
		})
		.join('');
	return {
		xml: `${DIAGRAM_XML_PROLOG}<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${body}</dsp:spTree></dsp:drawing>`,
		modelIds,
	};
}

/** Shared generated shape properties; hosts own model-id association and package placement. */
export function diagramShapePropertiesXml(shape: DiagramDrawingShape): string {
	if (
		!Object.values(shape.frame).every(Number.isFinite) ||
		shape.frame.width < 0 ||
		shape.frame.height < 0
	)
		throw new Error('The generated diagram has invalid geometry.');
	for (const path of shape.paths ?? [])
		for (const command of path.commands)
			if (
				Object.values(command).some((value) => typeof value === 'number' && !Number.isFinite(value))
			)
				throw new Error('The generated diagram has an invalid custom path.');
	const fill = shape.fill ? drawingFillXml(shape.fill) : '';
	if (shape.fill && fill === undefined)
		throw new Error('The generated diagram has an unsupported fill.');
	const preset = normalizeStShapeType(shape.geometry);
	if (!shape.paths?.length && !preset)
		throw new Error(`Unsupported diagram geometry: ${shape.geometry}`);
	const geometry = shape.paths?.length
		? `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${shape.paths.map(pathXml).join('')}</a:pathLst></a:custGeom>`
		: `<a:prstGeom prst="${preset}">${diagramAvLstXml(shape.adjustments)}</a:prstGeom>`;
	const frame = shape.frame;
	const transform = `<a:xfrm${shape.rotation ? ` rot="${Math.round(shape.rotation * 60000)}"` : ''}${shape.flipHorizontal ? ' flipH="1"' : ''}${shape.flipVertical ? ' flipV="1"' : ''}><a:off x="${Math.round(frame.x)}" y="${Math.round(frame.y)}"/><a:ext cx="${Math.round(frame.width)}" cy="${Math.round(frame.height)}"/></a:xfrm>`;
	return `<dsp:spPr>${transform}${geometry}${fill}${shape.line ? drawingLineXml(shape.line) : ''}${shape.effects ? drawingEffectsXml(shape.effects) : ''}</dsp:spPr>`;
}
