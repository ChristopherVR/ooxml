import { XMLBuilder } from 'fast-xml-parser';
import { diagramShapePropertiesXml } from '../../../../diagram/write/drawing';
import { normalizeStShapeType } from '../../../../geometry/preset-geometry-names';
import { drawingFillXml } from '../../../../drawingml/write-fill';

/**
 * Fabricate the cached diagram DRAWING part (`ppt/diagrams/drawingN.xml`) for
 * an SDK-created SmartArt element from its in-memory `drawingShapes`.
 *
 * Without this part PowerPoint recomputes the diagram from the (deliberately
 * simplified) fabricated `layoutN.xml`, which renders every node as an
 * identical rounded rectangle: the "all shapes became the same default shape"
 * symptom. Emitting a cached `dsp:drawing` whose `dsp:sp` shapes each carry
 * their own `a:prstGeom` preserves the per-node geometry (pyramid trapezoids,
 * cycle ellipses, chevrons, ...) the viewer computed.
 *
 * Shape model ids reference `type="pres"` presentation points. Each point's
 * `presAssocID` links back to its semantic content node, matching the
 * association model used by PowerPoint-authored diagrams.
 */
import { EMU_PER_PX } from '../../constants';
import { customGeometryPathsToXml } from '../../geometry/custom-geometry';
import { stripXmlOrderMarkers } from '../../geometry/custom-geometry-command-order';
import type {
	PptxElement,
	PptxSmartArtDrawingShape,
	PptxSmartArtNode,
	ShapePptxElement,
} from '../../types';
import { smartArtElementStyleLabel } from '../../utils/smartart-element-style-label';
import { drawingShape3dXml } from './smartart-fabrication-3d';
import { XML_PROLOG, avLstXml, xmlEscape } from './smartart-fabrication-data';
import { smartArtDrawingFill, smartArtDrawingLine } from './smartart-fabrication-paint';
import type { SmartArtRunWriter } from './smartart-fabrication-text';
import { drawingTextBodyXml } from './smartart-fabrication-text';

/** Content type for the cached diagram drawing part. */
export const DIAGRAM_DRAWING_CONTENT_TYPE =
	'application/vnd.ms-office.drawingml.diagramDrawing+xml';

/** Relationship type linking a data part to its cached drawing part. */
export const DIAGRAM_DRAWING_REL_TYPE =
	'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';

const DSP_XMLNS =
	'xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"' +
	' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
	' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** The six theme accent colours SmartArt cycles through, one per node. */
const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
const geometryBuilder = new XMLBuilder({
	ignoreAttributes: false,
	attributeNamePrefix: '@_',
	suppressEmptyNode: true,
	format: false,
});

/** Round a pixel measurement to whole EMU. */
function toEmu(px: number): number {
	return Math.round(px * EMU_PER_PX);
}

/**
 * Resolve the presentation-point GUID for a drawing shape.
 *
 * Layout-engine shapes embed the node id in their `id` (`engine-<nodeId>`,
 * `reflow-<...>-<nodeId>`); parsed shapes may match a node id directly. Falls
 * back to positional pairing, then to a fresh GUID so a shape is never dropped.
 */
function resolveShapeModelId(
	shape: PptxSmartArtDrawingShape,
	index: number,
	nodes: PptxSmartArtNode[],
	presentationGuidByNodeId: Map<string, string>,
	allocateModelId?: () => string,
): string {
	for (const presentationId of presentationGuidByNodeId.values()) {
		if (shape.id === presentationId) {
			return presentationId;
		}
	}
	const direct = presentationGuidByNodeId.get(shape.id);
	if (direct) {
		return direct;
	}
	const matched = nodes.find(
		(node) => node.id && (shape.id === node.id || shape.id.endsWith(`-${node.id}`)),
	);
	if (matched?.id) {
		const guid = presentationGuidByNodeId.get(matched.id);
		if (guid) {
			return guid;
		}
	}
	const positional = nodes[index]?.id;
	if (positional) {
		const guid = presentationGuidByNodeId.get(positional);
		if (guid) {
			return guid;
		}
	}
	// A shape of no node (a process or cycle arrow between two nodes) gets a presentation point
	// of its own, so it is drawn rather than dropped.
	return allocateModelId?.() ?? '';
}

function textTransformXml(shape: PptxSmartArtDrawingShape): string {
	return (
		`<dsp:txXfrm>` +
		`<a:off x="${toEmu(shape.x)}" y="${toEmu(shape.y)}"/>` +
		`<a:ext cx="${toEmu(Math.max(shape.width, 1))}" cy="${toEmu(Math.max(shape.height, 1))}"/>` +
		`</dsp:txXfrm>`
	);
}

function styleXml(index: number): string {
	const accent = ACCENTS[index % ACCENTS.length];
	return (
		`<dsp:style>` +
		`<a:lnRef idx="2"><a:schemeClr val="${accent}"><a:shade val="50000"/></a:schemeClr></a:lnRef>` +
		`<a:fillRef idx="1"><a:schemeClr val="${accent}"/></a:fillRef>` +
		`<a:effectRef idx="0"><a:schemeClr val="${accent}"/></a:effectRef>` +
		`<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>` +
		`</dsp:style>`
	);
}

function shapePropsXml(shape: PptxSmartArtDrawingShape): string {
	const fill = smartArtDrawingFill(shape);
	const line = smartArtDrawingLine(shape);
	const paint = { ...(fill && { fill }), ...(line && { line }) };
	if (
		!shape.customGeometryPaths?.length &&
		!shape.scene3d &&
		!shape.shape3d &&
		shape.skewX === undefined &&
		shape.skewY === undefined &&
		normalizeStShapeType(shape.shapeType ?? 'rect')
	) {
		return diagramShapePropertiesXml({
			modelId: shape.id,
			has3d: false,
			geometry: shape.shapeType ?? 'rect',
			frame: {
				x: toEmu(shape.x),
				y: toEmu(shape.y),
				width: toEmu(Math.max(shape.width, 1)),
				height: toEmu(Math.max(shape.height, 1)),
			},
			...(shape.rotation ? { rotation: shape.rotation } : {}),
			...(shape.shapeAdjustments ? { adjustments: shape.shapeAdjustments } : {}),
			...paint,
		});
	}
	const rot = shape.rotation ? ` rot="${Math.round(shape.rotation * 60000)}"` : '';
	const skewX = shape.skewX !== undefined ? ` skewX="${Math.round(shape.skewX * 60000)}"` : '';
	const skewY = shape.skewY !== undefined ? ` skewY="${Math.round(shape.skewY * 60000)}"` : '';
	const xfrm =
		`<a:xfrm${rot}${skewX}${skewY}>` +
		`<a:off x="${toEmu(shape.x)}" y="${toEmu(shape.y)}"/>` +
		`<a:ext cx="${toEmu(Math.max(shape.width, 1))}" cy="${toEmu(Math.max(shape.height, 1))}"/>` +
		`</a:xfrm>`;
	const prst = shape.shapeType && shape.shapeType !== 'custom' ? shape.shapeType : 'rect';
	const geom =
		shape.customGeometryPaths && shape.customGeometryPaths.length > 0
			? stripXmlOrderMarkers(
					geometryBuilder.build({
						'a:custGeom': customGeometryPathsToXml(
							shape.customGeometryPaths,
							shape.customGeometryRawData,
							{
								adjustHandlesXY: shape.customGeometryAdjustHandlesXY,
								adjustHandlesPolar: shape.customGeometryAdjustHandlesPolar,
								connectionSites: shape.customGeometryConnectionSites,
								textRect: shape.customGeometryTextRect,
							},
						),
					}),
				)
			: `<a:prstGeom prst="${xmlEscape(prst)}">${avLstXml(shape.shapeAdjustments)}</a:prstGeom>`;
	const fillXml = fill ? drawingFillXml(fill) : '';
	const lineWidth = line?.widthEmu ? ` w="${line.widthEmu}"` : '';
	const lineXml = line?.fill ? `<a:ln${lineWidth}>${drawingFillXml(line.fill)}</a:ln>` : '';
	return `<dsp:spPr>${xfrm}${geom}${fillXml}${lineXml}${drawingShape3dXml(shape)}</dsp:spPr>`;
}

function shapeXml(
	shape: PptxSmartArtDrawingShape,
	index: number,
	nodes: PptxSmartArtNode[],
	presentationGuidByNodeId: Map<string, string>,
	writeRun?: SmartArtRunWriter,
	allocateModelId?: () => string,
): string {
	const modelId = resolveShapeModelId(
		shape,
		index,
		nodes,
		presentationGuidByNodeId,
		allocateModelId,
	);
	if (!modelId) {
		return '';
	}
	return (
		`<dsp:sp modelId="${modelId}">` +
		`<dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>` +
		`${shapePropsXml(shape)}${styleXml(index)}${drawingTextBodyXml(shape, writeRun)}${textTransformXml(shape)}</dsp:sp>`
	);
}

/**
 * Build the complete `drawingN.xml` payload for a fabricated SmartArt diagram.
 *
 * Returns `undefined` when there are no drawing shapes to cache (the caller
 * then omits the drawing part and lets PowerPoint recompute the layout).
 */
export function buildFabricatedDrawingXml(
	shapes: PptxSmartArtDrawingShape[] | undefined,
	nodes: PptxSmartArtNode[],
	presentationGuidByNodeId: Map<string, string>,
	writeRun?: SmartArtRunWriter,
	/**
	 * Gives a shape that belongs to no node its presentation point id (in shape order); the
	 * caller writes those points (`transitionPresentationGuids`). Without it such a shape is left
	 * out of the drawing.
	 */
	allocateModelId?: () => string,
): string | undefined {
	if (!shapes || shapes.length === 0) {
		return undefined;
	}
	const body = shapes
		.map((shape, index) =>
			shapeXml(shape, index, nodes, presentationGuidByNodeId, writeRun, allocateModelId),
		)
		.join('');
	if (!body) {
		return undefined;
	}
	return (
		`${XML_PROLOG}\r\n<dsp:drawing ${DSP_XMLNS}>` +
		`<dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${body}</dsp:spTree></dsp:drawing>`
	);
}

/**
 * Convert decomposed SmartArt shape elements (from `decomposeSmartArt`) into
 * cacheable drawing shapes.
 *
 * When an SDK-created diagram carries no `drawingShapes`, the viewer renders it
 * by running the decompose/layout algorithms; the same algorithm output is
 * cached here so PowerPoint reopens the diagram with matching per-node geometry
 * instead of recomputing the simplified fabricated layout. Non-shape elements
 * (connectors) are skipped: they are reconstructed by PowerPoint's own layout.
 */
function withStyleLabel(element: PptxElement): { styleLabel?: string } {
	const styleLabel = smartArtElementStyleLabel(element);
	return styleLabel ? { styleLabel } : {};
}

export function smartArtElementsToDrawingShapes(
	elements: PptxElement[] | undefined,
): PptxSmartArtDrawingShape[] {
	if (!elements || elements.length === 0) {
		return [];
	}
	const shapes: PptxSmartArtDrawingShape[] = [];
	for (const el of elements) {
		if (el.type !== 'shape') {
			continue;
		}
		const shape = el as ShapePptxElement;
		shapes.push({
			id: shape.id,
			shapeType: shape.shapeType ?? 'rect',
			x: shape.x,
			y: shape.y,
			width: shape.width,
			height: shape.height,
			rotation: shape.rotation,
			skewX: shape.skewX,
			skewY: shape.skewY,
			...(shape.pathData || shape.customGeometryPaths?.length
				? {
						pathData: shape.pathData,
						pathWidth: shape.pathWidth,
						pathHeight: shape.pathHeight,
						customGeometryPaths: shape.customGeometryPaths,
						customGeometryRawData: shape.customGeometryRawData,
						customGeometryAdjustHandlesXY: shape.customGeometryAdjustHandlesXY,
						customGeometryAdjustHandlesPolar: shape.customGeometryAdjustHandlesPolar,
						customGeometryConnectionSites: shape.customGeometryConnectionSites,
						customGeometryTextRect: shape.customGeometryTextRect,
					}
				: {}),
			fillColor: shape.shapeStyle?.fillColor,
			...(shape.shapeStyle?.fillOpacity !== undefined && shape.shapeStyle.fillOpacity < 1
				? { fillOpacity: shape.shapeStyle.fillOpacity }
				: {}),
			...(shape.shapeAdjustments ? { shapeAdjustments: shape.shapeAdjustments } : {}),
			strokeColor: shape.shapeStyle?.strokeColor,
			strokeWidth: shape.shapeStyle?.strokeWidth,
			text: shape.text,
			...(shape.textSegments ? { textSegments: shape.textSegments } : {}),
			fontSize: shape.textStyle?.fontSize,
			fontColor: shape.textStyle?.color,
			...(shape.shapeStyle?.scene3d ? { scene3d: shape.shapeStyle.scene3d } : {}),
			...(shape.shapeStyle?.shape3d ? { shape3d: shape.shapeStyle.shape3d } : {}),
			...(shape.textStyle?.text3d ? { text3d: shape.textStyle.text3d } : {}),
			...withStyleLabel(el),
		});
	}
	return shapes;
}
