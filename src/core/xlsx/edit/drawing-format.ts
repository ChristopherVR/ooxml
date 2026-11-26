import { lookupPresetShape } from '../../geometry/preset-shape-evaluator';
import { inspectRasterImage } from '../../drawingml/raster-image';
import type { DrawingColor, DrawingFill, DrawingLine } from '../../drawingml/types';
import type { DrawingOuterShadow } from '../../drawingml/effect-model';
import { wordArtRgb } from '../../drawingml/word-art';
import { NS, buildXml, elements, parseXml } from '../../xml/index';
import { anchorToPixelBox, pixelBoxToAnchor, pixelSizeToExtent } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import type { ImageObject } from '../model';
import { selfContainedXml } from '../read/xml-util';
import {
	applyDrawingDescription,
	applyDrawingFrameFormat,
	type DrawingFrameFormat,
} from '../write/drawing-format';
import { nativeShapeXml } from '../write/native-shape';
import { newMediaPart } from './charts';
import { type EditContext, sheetAt, sheetDrawingAt } from './context';
import { setDrawingAnchor } from './drawings';
import { editDrawings, isWorksheetShape } from './native-shapes';

type Ink = string | { accent: 1 | 2 | 3 | 4 | 5 | 6 };
export type DrawingFormatIntent =
	| { kind: 'altText'; text: string }
	| { kind: 'size'; width: number; height: number }
	| { kind: 'geometry'; preset: string }
	| { kind: 'fill' | 'outline'; color: Ink | null }
	| { kind: 'shadow'; enabled: boolean }
	| { kind: 'shapeStyle'; color: Ink; outlined: boolean }
	| { kind: 'pictureStyle'; border: Ink | null; shadow: boolean }
	| { kind: 'resetPicture' }
	| { kind: 'replacePicture' | 'compressPicture'; bytes: Uint8Array };

function color(ink: Ink): DrawingColor {
	if (typeof ink !== 'string')
		return { kind: 'scheme', value: 'accent' + ink.accent, transforms: [] };
	const rgb = wordArtRgb(ink);
	if (!rgb) throw new RangeError('Invalid drawing color');
	return { kind: 'srgb', value: rgb, transforms: [] };
}

const fill = (ink: Ink | null): DrawingFill =>
	ink === null ? { kind: 'none' } : { kind: 'solid', color: color(ink) };
const line = (before: DrawingLine | undefined, ink: Ink | null, width = 12700): DrawingLine => ({
	...before,
	widthEmu: before?.widthEmu || width,
	fill: fill(ink),
});
const shadow = (): DrawingOuterShadow => ({
	kind: 'outerShadow',
	distanceEmu: 38100,
	directionAngle: 2700000,
	blurRadiusEmu: 50800,
	color: { kind: 'srgb', value: '000000', transforms: [{ name: 'alpha', value: '40000' }] },
});

function frameFormat(
	intent: DrawingFormatIntent,
	outline: DrawingLine | undefined,
): DrawingFrameFormat | undefined {
	switch (intent.kind) {
		case 'geometry':
			if (!lookupPresetShape(intent.preset)) throw new RangeError('Unknown shape preset');
			return { geometry: intent.preset };
		case 'fill':
			return { fill: fill(intent.color) };
		case 'outline':
			return { line: line(outline, intent.color) };
		case 'shadow':
			return { shadow: intent.enabled ? shadow() : null };
		case 'shapeStyle':
			return {
				fill: fill(intent.outlined ? null : intent.color),
				line: line(undefined, intent.color),
			};
		case 'pictureStyle':
			return {
				line: line(undefined, intent.border, 28575),
				shadow: intent.shadow ? shadow() : null,
			};
		case 'resetPicture':
			return { geometry: 'rect', line: line(undefined, null), shadow: null };
		default:
			return undefined;
	}
}

function applyImageFormat(image: ImageObject, format: DrawingFrameFormat): void {
	image.frameFormat = { ...image.frameFormat, ...format };
	if (format.line) image.line = format.line;
	if (format.shadow) image.shadow = format.shadow;
	else if (format.shadow === null) delete image.shadow;
}

export function formatDrawing(
	ctx: EditContext,
	sheet: number,
	index: number,
	intent: DrawingFormatIntent,
): boolean {
	const drawing = sheetDrawingAt(ctx.workbook, sheet, index);
	const shape = isWorksheetShape(drawing) ? drawing.shape : undefined;
	const image = drawing?.kind === 'image' ? drawing : undefined;
	if (!drawing || (!shape && !image)) return false;

	if (intent.kind === 'size') {
		const { width, height } = intent;
		if (![width, height].every((n) => Number.isFinite(n) && n >= 4 && n <= 10000))
			throw new RangeError('Drawing edges must be between 4 and 10000 pixels');
		const ws = sheetAt(ctx.workbook, sheet);
		const metrics = createGridMetrics(ws, { zoom: 100 });
		const box = anchorToPixelBox(ws, drawing.anchor, metrics);
		const anchor = drawing.anchor.to
			? pixelBoxToAnchor(ws, { ...box, w: width, h: height }, metrics, drawing.anchor)
			: { from: drawing.anchor.from, ext: pixelSizeToExtent(width, height) };
		setDrawingAnchor(ctx, sheet, index, anchor);
		return true;
	}

	if (intent.kind === 'replacePicture' || intent.kind === 'compressPicture') {
		if (!image || image.svgPartName || image.effects?.some((effect) => effect.kind === 'artistic'))
			return false;
		const { mimeType } = inspectRasterImage(intent.bytes);
		const workbook = ctx.workbook;
		const part = newMediaPart(workbook, mimeType);
		return editDrawings(
			ctx,
			sheet,
			intent.kind === 'compressPicture' ? 'Compress picture' : 'Change picture',
			() => {
				workbook.source ??= { parts: new Map() };
				workbook.source.parts.set(part, intent.bytes.slice());
				image.partName = part;
				image.contentType = mimeType;
				return true;
			},
		);
	}

	const format = frameFormat(intent, image?.line ?? shape?.line);
	if (image) {
		if (
			intent.kind === 'geometry' ||
			intent.kind === 'fill' ||
			intent.kind === 'shapeStyle' ||
			(image.effectDag && format?.shadow !== undefined)
		)
			return false;
		return editDrawings(ctx, sheet, 'Format picture', () => {
			if (format) applyImageFormat(image, format);
			if (intent.kind === 'altText') image.description = intent.text;
			if (intent.kind === 'resetPicture') {
				image.crop = {};
				image.layout = {};
				image.effects = image.effects?.filter((effect) => effect.kind === 'unsupported') ?? [];
			}
			return true;
		});
	}
	if (!shape || drawing.kind !== 'unsupported') return false;
	if (
		shape.equationParagraphs ||
		(format && shape.kind === 'group') ||
		(intent.kind === 'geometry' && (shape.kind !== 'shape' || shape.textBox)) ||
		intent.kind === 'pictureStyle' ||
		intent.kind === 'resetPicture' ||
		(shape.effectDag && format?.shadow !== undefined)
	)
		return false;
	const source =
		drawing.sourceXml ??
		shape.sourceShapeXml ??
		`<root xmlns:a="${NS.a}" xmlns:xdr="${NS.xdr}">${nativeShapeXml(shape, () => 1)}</root>`;
	const doc = parseXml(source);
	const root = doc.documentElement;
	const node =
		shape.sourceShapeXml && !drawing.sourceXml
			? root
			: elements(root).find(
					(item) =>
						item.namespaceURI === NS.xdr && ['sp', 'cxnSp', 'grpSp'].includes(item.localName),
				);
	if (!node) return false;
	if (intent.kind === 'altText') applyDrawingDescription(node, intent.text);
	if (format) applyDrawingFrameFormat(node, format);
	return editDrawings(ctx, sheet, 'Format shape', () => {
		if (drawing.sourceXml) drawing.sourceXml = buildXml(doc);
		else if (shape.kind !== 'group') shape.sourceShapeXml = selfContainedXml(node);
		if (intent.kind === 'altText') shape.description = intent.text;
		if (format?.geometry !== undefined) {
			shape.geometry = format.geometry;
			delete shape.paths;
			delete shape.adjustments;
		}
		if (format?.line) shape.line = format.line;
		if (format?.fill) shape.fill = format.fill;
		if (format?.shadow) shape.shadow = format.shadow;
		else if (format?.shadow === null) delete shape.shadow;
		return true;
	});
}
