// Resolves the shapes of a cached drawing part (`dsp:drawing`) into the laid-out view renderers
// and editing operations use (`DiagramShapeView`): pixels instead of EMU, `#RRGGBB` colours through
// a host-supplied resolver, and the text style of the first styled run. Generalises the pptx
// runtime's `parseDrawingShape` (`pptx/core/core/runtime/PptxHandlerRuntimeSmartArtParsing.ts`,
// `smartart-drawing-shape-style.ts`, `smartart-drawing-shape-font-ref.ts`, 3b2ecd9a6) without its
// theme and presentation dependencies; parity with it is pinned by
// `pptx/__tests__/integration/smartart-drawing-shape-neutral-parity.test.ts`.
import { EMU_PER_PIXEL } from '../units/constants';
import {
	resolveDrawingColor,
	type DrawingColorTheme,
	type ResolvedDrawingColor,
} from '../drawingml/drawing-color';
import { drawingParagraphText } from '../drawingml/read-text-body';
import type { DrawingRichTextBody } from '../drawingml/text-model';
import type { XmlDocument } from '../xml/index';
import { parseDiagramDrawing } from './drawing';
import type { DiagramShapeView } from './drawing-view';
import type { DiagramColor, DiagramDrawingShape, DiagramFill, DiagramIssue } from './types';

/** Resolves a DrawingML colour to hex and alpha; `undefined` when it cannot be determined. */
export type DiagramColorResolver = (color: DiagramColor) => ResolvedDrawingColor | undefined;

export interface DiagramShapeResolveOptions {
	resolveColor: DiagramColorResolver;
	/** EMU per output unit; defaults to EMU per CSS pixel (9525). */
	emuPerPx?: number;
	/** Label colour taken before the style matrix's `a:fontRef` when runs carry none. */
	textColor?: string;
}

/** A colour resolver over a host theme, using the shared DrawingML colour pipeline. */
export const themeColorResolver =
	(theme: DrawingColorTheme): DiagramColorResolver =>
	(color) =>
		resolveDrawingColor(color, theme);

const PX_PER_PT = 96 / 72;

type FillView = Pick<
	DiagramShapeView,
	| 'fillNone'
	| 'fillColor'
	| 'fillOpacity'
	| 'fillGradientStops'
	| 'fillGradientType'
	| 'fillGradientAngle'
	| 'fillPatternPreset'
	| 'fillPatternForegroundColor'
	| 'fillPatternBackgroundColor'
	| 'fillBlipEmbedId'
	| 'pictureFill'
>;

function resolveFill(fill: DiagramFill | undefined, resolve: DiagramColorResolver): FillView {
	const view: FillView = {};
	if (fill?.kind === 'picture') view.pictureFill = fill;
	const hex = (color: DiagramColor | undefined) => (color ? resolve(color)?.hex : undefined);
	switch (fill?.kind) {
		case 'none':
			view.fillNone = true;
			break;
		case 'solid': {
			const color = resolve(fill.color);
			if (color) view.fillColor = color.hex;
			if (color && color.alpha < 1) view.fillOpacity = color.alpha;
			break;
		}
		case 'gradient': {
			// Stops in position order; an opacity only where the colour declares an alpha transform.
			const stops = [...fill.stops]
				.sort((left, right) => left.position - right.position)
				.flatMap((stop) => {
					const color = resolve(stop.color);
					if (!color) return [];
					const alpha = stop.color.transforms.some((entry) => entry.name.startsWith('alpha'));
					return [
						{
							color: color.hex,
							position: stop.position,
							...(alpha ? { opacity: color.alpha } : {}),
						},
					];
				});
			if (stops.length === 0) break;
			view.fillGradientStops = stops;
			view.fillGradientType = fill.path ? 'radial' : 'linear';
			view.fillGradientAngle = fill.angle ?? 0;
			const middle = stops[Math.floor(stops.length / 2)]?.color;
			if (middle) view.fillColor = middle;
			break;
		}
		case 'pattern': {
			if (fill.preset.trim()) view.fillPatternPreset = fill.preset.trim();
			const foreground = hex(fill.foreground);
			const background = hex(fill.background);
			if (foreground) view.fillPatternForegroundColor = foreground;
			if (background) view.fillPatternBackgroundColor = background;
			const flat = foreground ?? background;
			if (flat) view.fillColor = flat;
			break;
		}
		case 'picture':
			if (fill.relId) view.fillBlipEmbedId = fill.relId;
			break;
		default:
			break;
	}
	return view;
}

type TextView = Pick<
	DiagramShapeView,
	| 'text'
	| 'fontSize'
	| 'fontColor'
	| 'fontFamily'
	| 'fontWeight'
	| 'fontStyle'
	| 'lineHeight'
	| 'lineHeightRatio'
	| 'lineSpacingAfter'
	| 'lineSpacingAfterRatio'
	| 'textInsetLeft'
	| 'textInsetTop'
	| 'textInsetRight'
	| 'textInsetBottom'
	| 'textVerticalAnchor'
>;

/** The text and the style of the first styled run, as the pptx cached-shape reader reports it. */
function resolveText(
	body: DrawingRichTextBody,
	resolve: DiagramColorResolver,
	emuPerPx: number,
): TextView {
	const view: TextView = { text: body.paragraphs.map(drawingParagraphText).join('\n') };
	for (const paragraph of body.paragraphs) {
		const lineSpacing = paragraph.properties?.lineSpacing;
		if (lineSpacing && lineSpacing.value > 0) {
			if (lineSpacing.unit === 'percent') view.lineHeightRatio ??= lineSpacing.value;
			else view.lineHeight ??= lineSpacing.value * PX_PER_PT;
		}
		const after = paragraph.properties?.spaceAfter;
		if (after && after.value >= 0) {
			if (after.unit === 'percent') view.lineSpacingAfterRatio ??= after.value;
			else view.lineSpacingAfter ??= after.value * PX_PER_PT;
		}
		for (const item of paragraph.items) {
			if (item.kind !== 'run' || !item.properties) continue;
			const run = item.properties;
			if (view.fontSize === undefined && run.sizePt !== undefined && run.sizePt > 0)
				view.fontSize = run.sizePt * PX_PER_PT;
			if (view.fontColor === undefined && run.fill?.kind === 'solid') {
				const color = resolve(run.fill.color)?.hex;
				if (color) view.fontColor = color;
			}
			const family = (
				run.eastAsian?.typeface ||
				run.latin?.typeface ||
				run.complexScript?.typeface ||
				''
			).trim();
			if (view.fontFamily === undefined && family) view.fontFamily = family;
			if (view.fontWeight === undefined && run.bold !== undefined)
				view.fontWeight = run.bold ? 700 : 400;
			if (view.fontStyle === undefined && run.italic !== undefined)
				view.fontStyle = run.italic ? 'italic' : 'normal';
			if (view.fontSize !== undefined && view.fontFamily && view.fontColor) break;
		}
		if (view.fontSize !== undefined && view.fontFamily && view.fontColor) break;
	}
	const properties = body.bodyProperties;
	const inset = (value: number | undefined) => (value === undefined ? undefined : value / emuPerPx);
	const insets = {
		textInsetLeft: inset(properties?.insetLeftEmu),
		textInsetTop: inset(properties?.insetTopEmu),
		textInsetRight: inset(properties?.insetRightEmu),
		textInsetBottom: inset(properties?.insetBottomEmu),
	};
	for (const [key, value] of Object.entries(insets))
		if (value !== undefined) view[key as keyof typeof insets] = value;
	const anchor = properties?.anchor?.trim();
	if (anchor) view.textVerticalAnchor = anchor;
	return view;
}

/** The laid-out view of one cached drawing shape. */
export function resolveDiagramDrawingShape(
	shape: DiagramDrawingShape,
	options: DiagramShapeResolveOptions,
): DiagramShapeView {
	const emuPerPx = options.emuPerPx ?? EMU_PER_PIXEL;
	const resolve = options.resolveColor;
	const px = (value: number) => Math.round(value / emuPerPx);
	const view: DiagramShapeView = {
		id: shape.modelId,
		shapeType: shape.geometry,
		x: px(shape.frame.x),
		y: px(shape.frame.y),
		width: px(shape.frame.width),
		height: px(shape.frame.height),
	};
	if (shape.rotation !== undefined) view.rotation = shape.rotation;
	if (shape.flipHorizontal) view.flipHorizontal = true;
	if (shape.flipVertical) view.flipVertical = true;
	if (shape.adjustments) view.shapeAdjustments = { ...shape.adjustments };
	Object.assign(view, resolveFill(shape.fill, resolve));
	const shadow = shape.shadowColor && resolve(shape.shadowColor)?.hex;
	if (shadow) {
		view.hasShadow = true;
		view.shadowColor = shadow;
	}
	const stroke = shape.line?.fill?.kind === 'solid' ? resolve(shape.line.fill.color) : undefined;
	if (stroke) view.strokeColor = stroke.hex;
	const width = shape.line?.widthEmu;
	if (width !== undefined && width > 0) view.strokeWidth = width / 12700;
	if (shape.body) Object.assign(view, resolveText(shape.body, resolve, emuPerPx));
	if (view.fontColor === undefined) {
		const fontRef = shape.style?.font?.color;
		const color = options.textColor ?? (fontRef ? resolve(fontRef)?.hex : undefined);
		if (color) view.fontColor = color;
	}
	if (shape.textFrame) {
		view.textFrameX = shape.textFrame.x / emuPerPx;
		view.textFrameY = shape.textFrame.y / emuPerPx;
		view.textFrameWidth = shape.textFrame.width / emuPerPx;
		view.textFrameHeight = shape.textFrame.height / emuPerPx;
	}
	return view;
}

/**
 * Whether a view keeps a drawable frame: a frame that rounds to zero width or height in the output
 * unit is a stale producer extent (or a hairline below half a pixel), except for `line` geometry,
 * whose rails and stems are cached with a zero height or width by design. The pptx reader applies
 * the same rule to its pixel frames.
 */
export const isDrawableDiagramShapeView = (view: DiagramShapeView): boolean =>
	view.shapeType === 'line' || (view.width > 0 && view.height > 0);

/**
 * Reads a drawing part straight into laid-out views: the DOM reader (`parseDiagramDrawing`) with
 * the host's colour resolver, without the frames `isDrawableDiagramShapeView` rejects. Issues the
 * reader reports are returned alongside.
 */
export function readDiagramDrawingShapes(
	source: string | XmlDocument,
	options: DiagramShapeResolveOptions,
	part?: string,
): { shapes: DiagramShapeView[]; issues: DiagramIssue[] } {
	const drawing = parseDiagramDrawing(source, part);
	return {
		shapes: drawing.shapes
			.map((shape) => resolveDiagramDrawingShape(shape, options))
			.filter(isDrawableDiagramShapeView),
		issues: drawing.issues,
	};
}
