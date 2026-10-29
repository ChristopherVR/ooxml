/**
 * Pure element creation functions for the headless PPTX SDK.
 *
 * Each function creates a valid {@link PptxElement} object ready to be
 * pushed into a slide's `elements` array. These are framework-agnostic
 * data constructors with no XML or ZIP involvement.
 *
 * @module sdk/ElementFactory
 */

import { DEFAULT_POWERPOINT_TABLE_STYLE_ID } from '../../core/runtime/table-style-defaults';
import { svgToCustomGeometryPaths } from '../../geometry/custom-geometry';
import type { PptxChartData, PptxChartSeries, PptxChartType } from '../../types/chart';
import type {
	TextPptxElement,
	ShapePptxElement,
	ConnectorPptxElement,
	ImagePptxElement,
	TablePptxElement,
	ChartPptxElement,
	MediaPptxElement,
	GroupPptxElement,
	PptxElement,
} from '../../types/elements';
import type { ShapeStyle } from '../../types/shape-style';
import type { PptxTableData, PptxTableRow, PptxTableCell } from '../../types/table';
import type { TextStyle, TextSegment } from '../../types/text';
import { chartSeriesGradientFromInput } from './chart-gradient-input';
import { applyChartInputLayout } from './chart-layout-operations';
import {
	buildShapeStyle,
	connectorLineEndOverride,
	mapStrokeToShapeStyle,
} from './shape-style-input';
import type { ShapeStyleInputs } from './shape-style-input';
import type {
	TextOptions,
	TextSegmentInput,
	ShapeOptions,
	ImageOptions,
	TableInput,
	TableOptions,
	ChartInput,
	ChartOptions,
	ConnectorOptions,
	MediaOptions,
	GroupOptions,
	TextStyleInput,
} from './types';

// ---------------------------------------------------------------------------
// ID generation
// ---------------------------------------------------------------------------

let idCounter = 0;

function generateId(prefix: string): string {
	idCounter += 1;
	return `${prefix}_${Date.now().toString(36)}_${idCounter}`;
}

/** Reset the ID counter (useful for tests). */
export function resetIdCounter(): void {
	idCounter = 0;
}

// ---------------------------------------------------------------------------
// Internal mapping helpers
// ---------------------------------------------------------------------------

/** The fill, stroke and effect inputs of a text box or shape, with `fill` resolved. */
function effectInputs(
	options: TextOptions | ShapeOptions | undefined,
	fill: ShapeStyleInputs['fill'],
): ShapeStyleInputs {
	return {
		fill,
		stroke: options?.stroke,
		shadow: options?.shadow,
		innerShadow: options?.innerShadow,
		glow: options?.glow,
		softEdge: options?.softEdge,
		reflection: options?.reflection,
	};
}

function mapTextStyleInput(opts?: Partial<TextStyleInput>): Partial<TextStyle> {
	if (!opts) {
		return {};
	}
	const ts: Partial<TextStyle> = {};
	if (opts.fontSize !== undefined) {
		ts.fontSize = opts.fontSize;
	}
	if (opts.fontFamily !== undefined) {
		ts.fontFamily = opts.fontFamily;
	}
	if (opts.bold !== undefined) {
		ts.bold = opts.bold;
	}
	if (opts.italic !== undefined) {
		ts.italic = opts.italic;
	}
	if (opts.underline !== undefined) {
		ts.underline = opts.underline;
	}
	if (opts.strikethrough !== undefined) {
		ts.strikethrough = opts.strikethrough;
	}
	if (opts.color !== undefined) {
		ts.color = opts.color;
	}
	if (opts.themeColorRef !== undefined) {
		ts.colorRef = opts.themeColorRef;
	}
	if (opts.alignment !== undefined) {
		ts.align = opts.alignment;
	}
	if (opts.verticalAlignment !== undefined) {
		ts.vAlign = opts.verticalAlignment;
	}
	if (opts.lineSpacing !== undefined) {
		ts.lineSpacing = opts.lineSpacing;
	}
	if (opts.spaceBefore !== undefined) {
		ts.paragraphSpacingBefore = opts.spaceBefore;
	}
	if (opts.spaceAfter !== undefined) {
		ts.paragraphSpacingAfter = opts.spaceAfter;
	}
	return ts;
}

function buildTextSegments(
	input: string | TextSegmentInput[],
	baseStyle?: Partial<TextStyleInput>,
): TextSegment[] {
	const base = mapTextStyleInput(baseStyle);
	if (typeof input === 'string') {
		const lines = input.split('\n');
		const segments: TextSegment[] = [];
		for (let i = 0; i < lines.length; i++) {
			segments.push({ text: lines[i], style: base as TextStyle });
			if (i < lines.length - 1) {
				segments.push({
					text: '\n',
					isParagraphBreak: true,
					style: base as TextStyle,
				});
			}
		}
		return segments;
	}
	return input.map((seg) => ({
		text: seg.text,
		style: { ...base, ...mapTextStyleInput(seg.style) } as TextStyle,
	}));
}

// ---------------------------------------------------------------------------
// Default positions
// ---------------------------------------------------------------------------

const DEFAULTS = {
	text: { x: 100, y: 100, width: 600, height: 50 },
	shape: { x: 200, y: 200, width: 300, height: 200 },
	image: { x: 100, y: 100, width: 400, height: 300 },
	table: { x: 50, y: 150, width: 860, height: 200 },
	chart: { x: 100, y: 150, width: 600, height: 400 },
	connector: { x: 100, y: 100, width: 200, height: 0 },
	media: { x: 100, y: 100, width: 480, height: 270 },
	group: { x: 0, y: 0, width: 600, height: 400 },
} as const;

function pos(
	type: keyof typeof DEFAULTS,
	opts?: Partial<{ x: number; y: number; width: number; height: number }>,
) {
	const d = DEFAULTS[type];
	return {
		x: opts?.x ?? d.x,
		y: opts?.y ?? d.y,
		width: opts?.width ?? d.width,
		height: opts?.height ?? d.height,
	};
}

// ---------------------------------------------------------------------------
// Element creation functions
// ---------------------------------------------------------------------------

/**
 * Create a text box element.
 *
 * @param text - Plain string or rich text segments.
 * @param options - Position, size, and text styling.
 * @returns A valid {@link TextPptxElement}.
 *
 * @example
 * ```ts
 * const el = createTextElement("Hello World", { fontSize: 36, bold: true });
 * ```
 */
export function createTextElement(
	text: string | TextSegmentInput[],
	options?: TextOptions,
): TextPptxElement {
	const p = pos('text', options);
	const plainText = typeof text === 'string' ? text : text.map((s) => s.text).join('');
	const textStyleBase: Partial<TextStyleInput> = {
		fontSize: options?.fontSize,
		fontFamily: options?.fontFamily,
		bold: options?.bold,
		italic: options?.italic,
		underline: options?.underline,
		strikethrough: options?.strikethrough,
		color: options?.color,
		alignment: options?.alignment,
		verticalAlignment: options?.verticalAlignment,
		lineSpacing: options?.lineSpacing,
	};

	return {
		type: 'text',
		id: generateId('txt'),
		...p,
		rotation: options?.rotation,
		opacity: options?.opacity,
		text: plainText,
		textStyle: mapTextStyleInput(textStyleBase) as TextStyle,
		textSegments: buildTextSegments(text, textStyleBase),
		shapeStyle: buildShapeStyle(effectInputs(options, options?.fill)),
	};
}

/**
 * Create a shape element with optional text overlay.
 *
 * @param shapeType - Preset geometry name (e.g. "rect", "ellipse", "roundRect").
 * @param options - Position, styling, and optional text.
 * @returns A valid {@link ShapePptxElement}.
 *
 * @example
 * ```ts
 * const el = createShapeElement("roundRect", {
 *   fill: { type: "solid", color: "#4472C4" },
 *   text: "Click me",
 * });
 * ```
 */
export function createShapeElement(shapeType: string, options?: ShapeOptions): ShapePptxElement {
	const p = pos('shape', options);
	const el: ShapePptxElement = {
		type: 'shape',
		id: generateId('shp'),
		...p,
		rotation: options?.rotation,
		opacity: options?.opacity,
		shapeType,
		shapeAdjustments: options?.adjustments,
		shapeStyle: buildShapeStyle(
			effectInputs(options, options?.fill ?? { type: 'solid', color: '#4472C4' }),
		),
	};
	if (options?.text) {
		el.text = options.text;
		el.textStyle = mapTextStyleInput(options.textStyle) as TextStyle;
		el.textSegments = buildTextSegments(options.text, options.textStyle);
	}
	return el;
}

/**
 * Create a connector (line) element.
 *
 * @param options - Connector type, stroke, arrows, and connection endpoints.
 * @returns A valid {@link ConnectorPptxElement}.
 */
export function createConnectorElement(options?: ConnectorOptions): ConnectorPptxElement {
	const p = pos('connector', options);
	const connType =
		options?.type === 'bent'
			? 'bentConnector3'
			: options?.type === 'curved'
				? 'curvedConnector3'
				: 'straightConnector1';

	// The SDK's own connector defaults (black, 1 px) apply field by field, so a
	// stroke that sets only a dash or an arrow does not fall through to the
	// connector writer's UI defaults instead.
	const strokeInput = options?.stroke;
	const stroke = mapStrokeToShapeStyle({
		...strokeInput,
		color: strokeInput?.color ?? '#000000',
		width: strokeInput?.width ?? 1,
	});
	const style: ShapeStyle = {
		...stroke,
		...connectorLineEndOverride('Start', options?.startArrow),
		...connectorLineEndOverride('End', options?.endArrow),
	} as ShapeStyle;
	style.connectorStartArrow ??= 'none';
	style.connectorEndArrow ??= 'none';

	if (options?.from) {
		style.connectorStartConnection = {
			shapeId: options.from.elementId,
			connectionSiteIndex: options.from.site,
		};
	}
	if (options?.to) {
		style.connectorEndConnection = {
			shapeId: options.to.elementId,
			connectionSiteIndex: options.to.site,
		};
	}

	return {
		type: 'connector',
		id: generateId('cxn'),
		...p,
		rotation: options?.rotation,
		shapeType: connType,
		shapeStyle: style,
	};
}

/**
 * Create an image element.
 *
 * @param source - Data URL (`data:image/...;base64,...`) or archive path.
 * @param options - Position, alt text, and crop settings.
 * @returns A valid {@link ImagePptxElement}.
 */
export function createImageElement(source: string, options?: ImageOptions): ImagePptxElement {
	const p = pos('image', options);
	const isDataUrl = source.startsWith('data:');
	return {
		type: 'image',
		id: generateId('img'),
		...p,
		rotation: options?.rotation,
		opacity: options?.opacity,
		imagePath: isDataUrl ? undefined : source,
		imageData: isDataUrl ? source : undefined,
		altText: options?.altText,
		cropLeft: options?.cropLeft,
		cropTop: options?.cropTop,
		cropRight: options?.cropRight,
		cropBottom: options?.cropBottom,
	};
}

/**
 * Create a table element.
 *
 * @param input - Table rows, columns, and styling options.
 * @param options - Position overrides.
 * @returns A valid {@link TablePptxElement}.
 *
 * @example
 * ```ts
 * const el = createTableElement({
 *   rows: [
 *     { cells: [{ text: "Name" }, { text: "Score" }] },
 *     { cells: [{ text: "Alice" }, { text: "95" }] },
 *   ],
 *   firstRow: true,
 *   bandRows: true,
 * });
 * ```
 */
export function createTableElement(input: TableInput, options?: TableOptions): TablePptxElement {
	const numCols = Math.max(...input.rows.map((r) => r.cells.length), 1);
	const rowHeight = input.rows.length > 0 ? Math.max(30, Math.floor(200 / input.rows.length)) : 30;
	const totalHeight = options?.height ?? Math.max(100, input.rows.length * rowHeight);
	const p = pos('table', { ...options, height: totalHeight });

	// Normalize column widths to proportions summing to 1
	let colWidths: number[];
	if (input.columnWidths && input.columnWidths.length === numCols) {
		const sum = input.columnWidths.reduce((a, b) => a + b, 0);
		colWidths = input.columnWidths.map((w) => w / sum);
	} else {
		colWidths = Array(numCols).fill(1 / numCols);
	}

	const rows: PptxTableRow[] = input.rows.map((rowInput) => ({
		height: rowInput.height ?? rowHeight,
		cells: rowInput.cells.map((cellInput): PptxTableCell => ({
			text: cellInput.text,
			gridSpan: cellInput.gridSpan,
			rowSpan: cellInput.rowSpan,
			style: cellInput.style
				? {
						fontSize: cellInput.style.fontSize,
						bold: cellInput.style.bold,
						italic: cellInput.style.italic,
						color: cellInput.style.color,
						align: cellInput.style.alignment,
					}
				: undefined,
		})),
	}));

	const tableData: PptxTableData = {
		rows,
		columnWidths: colWidths,
		bandedRows: input.bandRows,
		bandedColumns: input.bandColumns,
		firstRowHeader: input.firstRow,
		lastRow: input.lastRow,
		firstCol: input.firstCol,
		lastCol: input.lastCol,
		// Match PowerPoint's "Insert > Table": a new table gets Medium Style 2 -
		// Accent 1 unless the caller picked a style. This is decided at
		// creation time so the save pipeline never has to invent a style for
		// loaded tables that legitimately have none.
		tableStyleId: input.style ?? DEFAULT_POWERPOINT_TABLE_STYLE_ID,
	};

	return {
		type: 'table',
		id: generateId('tbl'),
		...p,
		tableData,
	};
}

/**
 * Create a chart element.
 *
 * @param chartType - One of the supported chart type discriminators.
 * @param input - Series data, categories, title, legend options.
 * @param options - Position overrides.
 * @returns A valid {@link ChartPptxElement}.
 *
 * @example
 * ```ts
 * const el = createChartElement("bar", {
 *   series: [{ name: "Q1", values: [45, 62, 38] }],
 *   categories: ["A", "B", "C"],
 *   title: "Sales",
 * });
 * ```
 */
export function createChartElement(
	chartType: PptxChartType,
	input: ChartInput,
	options?: ChartOptions,
): ChartPptxElement {
	const p = pos('chart', options);

	const series: PptxChartSeries[] = input.series.map((s, index) => ({
		name: s.name,
		values: s.values,
		color: s.color,
		...chartSeriesGradientFromInput(chartType, s.gradientFill, index),
		...(s.boxWhiskerOptions ? { boxWhiskerOptions: { ...s.boxWhiskerOptions } } : {}),
		...(s.histogramOptions ? { histogramOptions: { ...s.histogramOptions } } : {}),
		...(s.waterfallOptions
			? {
					waterfallOptions: {
						...s.waterfallOptions,
						...(s.waterfallOptions.subtotalIndices
							? { subtotalIndices: [...s.waterfallOptions.subtotalIndices] }
							: {}),
					},
				}
			: {}),
		...(s.regionMapOptions
			? {
					regionMapOptions: {
						...s.regionMapOptions,
						...(s.regionMapOptions.entityIds
							? { entityIds: [...s.regionMapOptions.entityIds] }
							: {}),
					},
				}
			: {}),
		...(s.treemapOptions ? { treemapOptions: { ...s.treemapOptions } } : {}),
	}));

	const chartData: PptxChartData = {
		chartType,
		categories: input.categories,
		categoryLevels: input.categoryLevels?.map((level) => [...level]),
		series,
		title: input.title,
		grouping: input.grouping,
		...(input.barDirection !== undefined ? { barDirection: input.barDirection } : {}),
		style: {
			hasLegend: input.hasLegend ?? true,
			legendPosition: input.legendPosition,
		},
	};

	const element: ChartPptxElement = {
		type: 'chart',
		id: generateId('cht'),
		...p,
		chartData,
	};
	applyChartInputLayout(element, input);
	return element;
}

/**
 * Create a media element (audio or video).
 *
 * @param mediaType - "video" or "audio".
 * @param source - Data URL or archive path for the media file.
 * @param options - Playback options and position.
 * @returns A valid {@link MediaPptxElement}.
 */
export function createMediaElement(
	mediaType: 'video' | 'audio',
	source: string,
	options?: MediaOptions,
): MediaPptxElement {
	const p = pos('media', options);
	const isDataUrl = source.startsWith('data:');
	return {
		type: 'media',
		id: generateId('med'),
		...p,
		rotation: options?.rotation,
		mediaType,
		mediaPath: isDataUrl ? undefined : source,
		mediaData: isDataUrl ? source : undefined,
		autoPlay: options?.autoPlay,
		loop: options?.loop,
		volume: options?.volume,
		trimStartMs: options?.trimStartMs,
		trimEndMs: options?.trimEndMs,
		posterFrameData: options?.posterFrame,
	};
}

/**
 * Create a group element containing child elements.
 *
 * @param children - Array of child elements.
 * @param options - Position overrides for the group container.
 * @returns A valid {@link GroupPptxElement}.
 */
export function createGroupElement(
	children: PptxElement[],
	options?: GroupOptions,
): GroupPptxElement {
	const p = pos('group', options);
	return {
		type: 'group',
		id: generateId('grp'),
		...p,
		rotation: options?.rotation,
		children,
	};
}

/**
 * Create a freeform shape element from SVG path data.
 *
 * Uses a custom geometry (a:custGeom) shape with the provided path data.
 * The path is parsed into structured {@link CustomGeometryPath} segments
 * for round-tripping through OOXML serialization.
 *
 * @param pathData - An SVG path data string (e.g. `"M 0 0 L 100 50 L 50 100 Z"`).
 * @param options - Optional position, styling, and size overrides.
 * @returns A valid {@link ShapePptxElement} with custom geometry.
 *
 * @example
 * ```ts
 * const el = createFreeformElement("M 0 0 C 33 0 66 100 100 100", {
 *   stroke: { color: "#FF0000", width: 2 },
 * });
 * ```
 */
export function createFreeformElement(pathData: string, options?: ShapeOptions): ShapePptxElement {
	const p = pos('shape', options);
	const pathWidth = options?.width ?? p.width;
	const pathHeight = options?.height ?? p.height;
	const customGeometryPaths = svgToCustomGeometryPaths(pathData, pathWidth, pathHeight);

	const el: ShapePptxElement = {
		type: 'shape',
		id: generateId('frm'),
		...p,
		rotation: options?.rotation,
		opacity: options?.opacity,
		shapeType: 'custom',
		pathData,
		pathWidth,
		pathHeight,
		customGeometryPaths,
		shapeStyle: buildShapeStyle({
			...effectInputs(options, options?.fill ?? { type: 'none' }),
			stroke: options?.stroke ?? { color: '#000000', width: 2 },
		}),
	};
	if (options?.text) {
		el.text = options.text;
		el.textStyle = mapTextStyleInput(options.textStyle) as TextStyle;
		el.textSegments = buildTextSegments(options.text, options.textStyle);
	}
	return el;
}
