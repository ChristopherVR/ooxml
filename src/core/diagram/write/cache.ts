import { NS, buildXml, children, first, parseXml, type XmlElement } from '../../xml/index';
import { descendants } from '../dom';
import { createOfficeGuid } from '../../crypto/uuid';
import { parseDrawingRichTextBody } from '../../drawingml/read-text-body';
import { resolveDrawingColor } from '../../drawingml/drawing-color';
import { resolveDrawingOuterShadow } from '../../drawingml/drawing-shadow';
import type { DrawingColor } from '../../drawingml/types';
import { parseDiagramDataModel } from '../data-model';
import { parseDiagramColors } from '../definitions';
import { diagramNodesInTreeOrder } from '../edit/node-tree';
import { previewDiagramLayout, type DiagramPreviewInput } from '../layout-preview';
import { resolveDiagramDrawingStyles } from '../drawing-theme';
import {
	GENERATED_DIAGRAM_CACHE_NS,
	generatedDrawingDigest,
	DEFAULT_GENERATED_DIAGRAM_SCHEME,
} from '../cache-metadata';
import { writeGeneratedDiagramDrawing } from './drawing';
import { withDiagramPictureSources } from '../picture-layout';

const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
function add(
	parent: XmlElement,
	namespace: string,
	name: string,
	attrs: Record<string, string> = {},
): XmlElement {
	const child = parent.ownerDocument.createElementNS(namespace, name);
	for (const [key, value] of Object.entries(attrs)) child.setAttribute(key, value);
	parent.appendChild(child);
	return child;
}

/** Add the host relationship id without touching other extension payloads. */
export function withDiagramDrawingRelationship(xml: string, relationshipId: string): string {
	const doc = parseXml(xml);
	const model = doc.documentElement;
	let extension = descendants(model, 'dataModelExt').find((node) => node.namespaceURI === DSP);
	if (!extension) {
		const extLst = first(model, 'extLst', NS.dgm) ?? add(model, NS.dgm, 'dgm:extLst');
		extension = add(add(extLst, NS.a, 'a:ext', { uri: DSP }), DSP, 'dsp:dataModelExt', {
			minVer: NS.dgm,
		});
	}
	extension.setAttribute('relId', relationshipId);
	return buildXml(doc);
}

/** Generate only an admitted 2D interpretation. A declined cache leaves the host's fallback intact. */
export function buildGeneratedDiagramCache(
	dataXml: string,
	input: Omit<DiagramPreviewInput, 'nodes'>,
):
	| {
			dataXml: string;
			drawingXml: string;
			pictureParts?: Record<string, string>;
	  }
	| undefined {
	dataXml = withDiagramPictureSources(dataXml, input.source);
	const source = { ...input.source, dataXml };
	const model = parseDiagramDataModel(dataXml);
	const scheme = { ...DEFAULT_GENERATED_DIAGRAM_SCHEME, ...input.scheme };
	const preview = previewDiagramLayout({
		...input,
		source,
		scheme,
		nodes: diagramNodesInTreeOrder(model),
	});
	if (preview.mode !== 'interpreted' || preview.issues.length || !preview.drawing?.shapes.length)
		return undefined;
	const drawing = resolveDiagramDrawingStyles(preview.drawing, input.formatScheme);
	const doc = parseXml(dataXml);
	const root = doc.documentElement;
	const points = first(root, 'ptLst', NS.dgm);
	const connections = first(root, 'cxnLst', NS.dgm);
	if (!points || !connections) return undefined;
	const colors = source.colorsXml ? parseDiagramColors(source.colorsXml).labels : [];
	const authoredColors = colors.flatMap((label) => [...label.fill.colors, ...label.line.colors]);
	const themed = new Map<string, DrawingColor>();
	for (const color of authoredColors) {
		const resolved = resolveDrawingColor(color, { scheme: (name) => scheme[name] });
		if (resolved && !resolved.unapplied.length) themed.set(resolved.hex.toUpperCase(), color);
	}
	for (const shape of drawing.shapes) {
		if (shape.has3d || shape.foldedNodeIds?.length) return undefined;
		if (shape.fill && !['none', 'solid', 'gradient', 'picture'].includes(shape.fill.kind))
			return undefined;
		if (
			shape.fill?.kind === 'picture' &&
			(!shape.fill.relId || !drawing.pictureParts?.[shape.fill.relId])
		)
			return undefined;
		if (
			shape.fill?.kind === 'gradient' &&
			shape.fill.path &&
			!['circle', 'rect'].includes(shape.fill.path)
		)
			return undefined;
		if (
			shape.effects?.effects.some(
				(effect) =>
					effect.kind !== 'outerShadow' ||
					!resolveDrawingOuterShadow(effect, { scheme: (name) => scheme[name] }),
			)
		)
			return undefined;
		const point = children(points, 'pt', NS.dgm).find(
			(point) => point.getAttribute('modelId') === shape.sourceNodeId,
		);
		const direct = model.nodes.find((node) => node.modelId === shape.sourceNodeId);
		const fillIndex = shape.style?.fill?.index;
		const template = fillIndex ? input.formatScheme?.fillStyles[fillIndex - 1]?.fill : undefined;
		if (
			!direct?.fill &&
			shape.fill?.kind === 'solid' &&
			shape.fill.color.kind === 'srgb' &&
			(!template ||
				(template.kind === 'solid' &&
					template.color.kind === 'scheme' &&
					template.color.value === 'phClr'))
		) {
			const original = themed.get(`#${shape.fill.color.value}`.toUpperCase());
			if (original)
				shape.fill.color = {
					...original,
					transforms: [...original.transforms, ...shape.fill.color.transforms],
				};
		}
		if (!direct?.fill && shape.fill?.kind === 'gradient' && template?.kind === 'gradient') {
			shape.fill.stops.forEach((stop, index) => {
				const authored = template.stops[index]?.color;
				if (
					authored?.kind === 'scheme' &&
					authored.value === 'phClr' &&
					stop.color.kind === 'srgb'
				) {
					const original = themed.get(`#${stop.color.value}`.toUpperCase());
					if (original)
						stop.color = {
							...original,
							transforms: [...original.transforms, ...stop.color.transforms],
						};
				}
			});
		}
		const role =
			colors.find((label) => label.name === shape.sourceStyleLabel) ??
			colors.find((label) => label.name === 'node0');
		const fallbackColor = role?.textFill.colors[0] ?? {
			kind: 'scheme' as const,
			value: 'tx1',
			transforms: [],
		};
		const body = parseDrawingRichTextBody(first(point, 't', NS.dgm));
		if (body) {
			const props = body.bodyProperties;
			if (
				props &&
				((props.vertical && props.vertical !== 'horz') ||
					(props.columnCount ?? 1) > 1 ||
					props.rotation ||
					props.extraXml?.length ||
					props.autofit?.kind === 'shape')
			)
				return undefined;
			for (const paragraph of body.paragraphs) {
				if (
					(paragraph.properties?.bullet && paragraph.properties.bullet.kind !== 'none') ||
					paragraph.properties?.tabStops?.length ||
					paragraph.items.some((item) => item.kind === 'unknown')
				)
					return undefined;
				if (
					paragraph.properties?.extraXml?.length ||
					paragraph.items.some((item) => {
						const properties = 'properties' in item ? item.properties : undefined;
						return (
							properties &&
							(properties.baseline ||
								properties.spacingPt ||
								properties.rightToLeft ||
								properties.highlight ||
								properties.line ||
								(properties.caps && properties.caps !== 'none') ||
								(properties.strike && properties.strike !== 'noStrike') ||
								properties.extraXml?.length)
						);
					})
				)
					return undefined;
				paragraph.properties = {
					...paragraph.properties,
					defaultRunProperties: {
						sizePt: shape.text?.paragraphs[0]?.runs[0]?.sizePt ?? 12,
						fill: { kind: 'solid', color: fallbackColor },
						latin: { typeface: '+mn-lt' },
						...paragraph.properties?.defaultRunProperties,
					},
				};
			}
			body.bodyProperties = { anchor: 'ctr', ...body.bodyProperties };
			shape.body = body;
		}
	}
	let serialized: ReturnType<typeof writeGeneratedDiagramDrawing>;
	try {
		serialized = writeGeneratedDiagramDrawing(drawing);
	} catch {
		return undefined;
	}
	const oldMarkers = descendants(root, 'generatedCache').filter(
		(node) => node.namespaceURI === GENERATED_DIAGRAM_CACHE_NS,
	);
	const oldIds = new Set(
		oldMarkers.flatMap((marker) =>
			descendants(marker, 'shape').map((node) => node.getAttribute('id')),
		),
	);
	for (const point of children(points, 'pt', NS.dgm))
		if (oldIds.has(point.getAttribute('modelId'))) points.removeChild(point);
	for (const edge of children(connections, 'cxn', NS.dgm))
		if (oldIds.has(edge.getAttribute('destId'))) connections.removeChild(edge);
	for (const marker of oldMarkers) {
		const parent = marker.parentNode as XmlElement | null;
		if (
			parent &&
			children(parent, 'generatedCache', GENERATED_DIAGRAM_CACHE_NS).length === 1 &&
			parent.childNodes.length === 1
		)
			parent.parentNode?.removeChild(parent);
		else parent?.removeChild(marker);
	}
	const extLst = first(root, 'extLst', NS.dgm) ?? add(root, NS.dgm, 'dgm:extLst');
	const marker = add(
		add(extLst, NS.a, 'a:ext', { uri: GENERATED_DIAGRAM_CACHE_NS }),
		GENERATED_DIAGRAM_CACHE_NS,
		'oc:generatedCache',
		{
			engine: 'ordered-2d',
			approximate: '1',
			drawingSha256: generatedDrawingDigest(serialized.xml),
		},
	);
	drawing.shapes.forEach((shape, index) => {
		const id = serialized.modelIds.get(shape.modelId)!;
		const point = add(points, NS.dgm, 'dgm:pt', { modelId: id, type: 'pres' });
		add(point, NS.dgm, 'dgm:prSet', {
			...(shape.sourceNodeId ? { presAssocID: shape.sourceNodeId } : {}),
			presName: shape.sourceLayoutRole ?? 'node',
			presStyleLbl: shape.sourceStyleLabel ?? 'node1',
			presStyleIdx: String(index),
			presStyleCnt: String(drawing.shapes.length),
		});
		const properties = add(point, NS.dgm, 'dgm:spPr');
		if (shape.fill?.kind === 'picture' && shape.fill.sourceXml) {
			properties.appendChild(doc.importNode(parseXml(shape.fill.sourceXml).documentElement, true));
		}
		if (shape.sourceNodeId)
			add(connections, NS.dgm, 'dgm:cxn', {
				modelId: createOfficeGuid(),
				type: 'presOf',
				srcId: shape.sourceNodeId,
				destId: id,
				srcOrd: '0',
				destOrd: '0',
				presId: input.layoutId ?? '',
			});
		add(marker, GENERATED_DIAGRAM_CACHE_NS, 'oc:shape', { id });
	});
	return {
		dataXml: buildXml(doc),
		drawingXml: serialized.xml,
		...(drawing.pictureParts ? { pictureParts: drawing.pictureParts } : {}),
	};
}
