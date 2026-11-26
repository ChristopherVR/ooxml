import { lookupPresetShape } from '../../geometry/preset-shape-evaluator';
import { inspectRasterImage } from '../../drawingml/raster-image';
import type { DrawingColor, DrawingFill, DrawingLine } from '../../drawingml/types';
import { PICTURE_STYLES } from '../../drawingml/picture-styles-catalog';
import { pictureStyleFormat } from '../../drawingml/picture-style-format';
import { DEFAULT_OUTER_SHADOW_PRESET } from '../../drawingml/shape-effects-catalog';
import {
	presetShapeStyle,
	themeShapeStyle,
	type ShapeStylePick,
} from '../../drawingml/shape-style-catalog';
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
	| ({ kind: 'shapeStyle' } & ShapeStylePick)
	| { kind: 'pictureStyle'; style: string }
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
const line = (before: DrawingLine | undefined, ink: Ink | null): DrawingLine => ({
	...before,
	widthEmu: before?.widthEmu || 12700,
	fill: fill(ink),
});
function shadow(): DrawingOuterShadow {
	const { distEmu, dir, blurEmu, alpha, algn } = DEFAULT_OUTER_SHADOW_PRESET;
	return {
		kind: 'outerShadow',
		distanceEmu: distEmu,
		directionAngle: dir,
		blurRadiusEmu: blurEmu,
		...(algn ? { alignment: algn } : {}),
		color: {
			kind: 'srgb',
			value: '000000',
			transforms: alpha === undefined ? [] : [{ name: 'alpha', value: String(alpha) }],
		},
	};
}

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
		case 'shapeStyle': {
			const { block, row, column } = intent;
			const themed = block === 'theme' ? themeShapeStyle(row, column) : undefined;
			const preset = block === 'preset' ? presetShapeStyle(row, column) : undefined;
			const base = { effects: null, scene3d: null, shape3d: null };
			if (themed) return { ...base, style: themed, fill: null, line: null };
			if (preset) return { ...base, ...preset };
			throw new RangeError('Unknown shape style');
		}
		case 'pictureStyle': {
			const style = PICTURE_STYLES.find(({ key }) => key === intent.style);
			if (!style) throw new RangeError('Unknown picture style');
			return pictureStyleFormat(style);
		}
		case 'resetPicture':
			return { line: null, fill: null, effects: null, scene3d: null, shape3d: null };
		default:
			return undefined;
	}
}

function updateFrameEffects(
	model: Pick<ImageObject, 'effectDag' | 'shadow' | 'hasExplicitEffects'>,
	format: DrawingFrameFormat,
): void {
	if (format.effects !== undefined) {
		delete model.effectDag;
		const outer = format.effects?.effects.find((effect) => effect.kind === 'outerShadow');
		if (outer) model.shadow = outer;
		else delete model.shadow;
		if (format.effects === null) delete model.hasExplicitEffects;
		else model.hasExplicitEffects = true;
	}
	if (format.shadow !== undefined) {
		model.hasExplicitEffects = true;
		if (format.shadow) model.shadow = format.shadow;
		else delete model.shadow;
	}
}

function applyImageFormat(image: ImageObject, format: DrawingFrameFormat): void {
	const merged = { ...image.frameFormat, ...format };
	image.frameFormat = merged;
	if (format.geometry !== undefined) {
		image.geometry = format.geometry;
		delete merged.adjustments;
		delete image.adjustments;
		if (format.adjustments) image.adjustments = merged.adjustments = format.adjustments;
	}
	if (format.line) image.line = format.line;
	else if (format.line === null) delete image.line;
	if (format.effects !== undefined) delete merged.shadow;
	updateFrameEffects(image, format);
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
		else if (format?.line === null) delete shape.line;
		if (format?.fill) shape.fill = format.fill;
		else if (format?.fill === null) delete shape.fill;
		if (format?.style) shape.style = format.style;
		if (format) updateFrameEffects(shape, format);
		return true;
	});
}
