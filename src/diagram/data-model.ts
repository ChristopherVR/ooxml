// DOM parser for the diagram data model part (`dgm:dataModel`). Generalises
// `pptx/core/builders/PptxSmartArtParser.ts` and `parseSmartArtConnections` of the pptx runtime.
import { parseXml, type XmlDocument } from '../xml/index.js';
import { parseConnectionAttributes, parseCustomLayoutAttributes } from './attributes.js';
import {
	NS,
	attributeReader,
	children,
	descendants,
	first,
	integerAttribute,
	stringAttribute,
	type XmlElement,
} from './dom.js';
import type { DiagramConnection, DiagramDataModel, DiagramIssue, DiagramPoint } from './types.js';

/** `parTrans` and `sibTrans` points carry the text drawn on an org-chart relationship line. */
const TRANSITION_TYPES: ReadonlySet<string> = new Set(['parTrans', 'sibTrans']);
/** Point types that are not user content. */
const STRUCTURAL_TYPES: ReadonlySet<string> = new Set(['doc', 'pres', 'parTrans', 'sibTrans']);

/**
 * `transitionPointModelId -> text` for every transition point that carries non-empty text.
 * Generic over the point representation so pptx can feed it its object-tree points.
 */
export function collectTransitionText<T>(
	points: Iterable<T>,
	describe: (point: T) => { type?: string; modelId?: string },
	text: (point: T) => string,
): Map<string, string> {
	const byId = new Map<string, string>();
	for (const point of points) {
		const { type, modelId } = describe(point);
		if (!TRANSITION_TYPES.has((type ?? '').trim())) continue;
		const id = (modelId ?? '').trim();
		if (!id) continue;
		const value = text(point).trim();
		if (value.length > 0) byId.set(id, value);
	}
	return byId;
}

/** The label of a connection: the parent transition's text, else the sibling transition's. */
export function connectionLabel(
	connection: DiagramConnection,
	transitionText: ReadonlyMap<string, string>,
): string | undefined {
	const parent = connection.parentTransitionId
		? transitionText.get(connection.parentTransitionId)
		: undefined;
	const sibling = connection.siblingTransitionId
		? transitionText.get(connection.siblingTransitionId)
		: undefined;
	return parent ?? sibling;
}

/** Paragraphs of a `dgm:t` / `dsp:txBody` text body: the concatenated `a:t` of each `a:p`. */
export function textBodyParagraphs(body: XmlElement | undefined): string[] {
	if (!body) return [];
	return children(body, 'p', NS.a).map((paragraph) =>
		descendants(paragraph, 't')
			.map((run) => run.textContent ?? '')
			.join(''),
	);
}

function parsePoint(element: XmlElement): DiagramPoint | undefined {
	const modelId = stringAttribute(element, 'modelId');
	if (!modelId) return undefined;
	const paragraphs = textBodyParagraphs(first(element, 't', NS.dgm));
	const prSet = first(element, 'prSet', NS.dgm);
	const point: DiagramPoint = {
		modelId,
		type: stringAttribute(element, 'type') ?? 'node',
		paragraphs,
		text: paragraphs.join('\n'),
	};
	const optional: [keyof DiagramPoint, string | number | undefined][] = [
		['connectionId', stringAttribute(element, 'cxnId')],
		['presAssocId', stringAttribute(prSet, 'presAssocID')],
		['presName', stringAttribute(prSet, 'presName')],
		['presStyleLabel', stringAttribute(prSet, 'presStyleLbl')],
		['presStyleIndex', integerAttribute(prSet, 'presStyleIdx')],
		['presStyleCount', integerAttribute(prSet, 'presStyleCnt')],
		['layoutTypeId', stringAttribute(prSet, 'loTypeId')],
		['layoutCategoryId', stringAttribute(prSet, 'loCatId')],
		['quickStyleTypeId', stringAttribute(prSet, 'qsTypeId')],
		['colorsTypeId', stringAttribute(prSet, 'csTypeId')],
	];
	for (const [key, value] of optional)
		if (value !== undefined) (point as unknown as Record<string, unknown>)[key] = value;
	const customLayout = prSet ? parseCustomLayoutAttributes(attributeReader(prSet)) : undefined;
	if (customLayout) point.customLayout = customLayout;
	return point;
}

/** Parses a data model part. Malformed XML throws; structural problems are returned as `issues`. */
export function parseDiagramDataModel(
	source: string | XmlDocument,
	part?: string,
): DiagramDataModel {
	const root = (typeof source === 'string' ? parseXml(source, { label: 'DiagramML' }) : source)
		.documentElement;
	const issues: DiagramIssue[] = [];
	const report = (code: string, message: string) =>
		issues.push({ code, message, ...(part ? { part } : {}) });

	const points: DiagramPoint[] = [];
	const ids = new Set<string>();
	for (const element of children(first(root, 'ptLst', NS.dgm) ?? root, 'pt', NS.dgm)) {
		const point = parsePoint(element);
		if (!point) {
			report('POINT_ID_REQUIRED', 'dgm:pt requires modelId.');
			continue;
		}
		if (ids.has(point.modelId))
			report('POINT_ID_DUPLICATE', `Duplicate dgm:pt modelId: ${point.modelId}.`);
		ids.add(point.modelId);
		points.push(point);
	}

	const transitionText = collectTransitionText(
		points,
		(point) => ({ type: point.type, modelId: point.modelId }),
		(point) => point.text,
	);
	const connections: DiagramConnection[] = [];
	const parentById = new Map<string, string>();
	const connectionIds = new Set<string>();
	for (const element of children(first(root, 'cxnLst', NS.dgm) ?? root, 'cxn', NS.dgm)) {
		const connection = parseConnectionAttributes(attributeReader(element));
		if (!connection) {
			report('CONNECTION_ATTRIBUTE_REQUIRED', 'dgm:cxn requires srcId and destId.');
			continue;
		}
		if (connection.modelId && connectionIds.has(connection.modelId))
			report('CONNECTION_ID_DUPLICATE', `Duplicate dgm:cxn modelId: ${connection.modelId}.`);
		if (connection.modelId) connectionIds.add(connection.modelId);
		for (const [attribute, id] of [
			['srcId', connection.sourceId],
			['destId', connection.destId],
		] as const)
			if (!ids.has(id))
				report(
					'CONNECTION_ENDPOINT_MISSING',
					`dgm:cxn ${attribute} references missing point: ${id}.`,
				);
		const label = connectionLabel(connection, transitionText);
		connections.push(label ? { ...connection, label } : connection);
		// `parOf` (the default when `type` is omitted) is the only edge that expresses parent/child.
		if ((!connection.type || connection.type === 'parOf') && !parentById.has(connection.destId))
			parentById.set(connection.destId, connection.sourceId);
	}

	const ext = descendants(root, 'dataModelExt')[0];
	const drawingRelId = stringAttribute(ext, 'relId');
	return {
		points,
		connections,
		nodes: points.filter((point) => !STRUCTURAL_TYPES.has(point.type)),
		parentById,
		...(drawingRelId ? { drawingRelId } : {}),
		issues,
	};
}
