import { parseConnectionAttributes, parseCustomLayoutAttributes } from '../../../diagram/index.js';
import type {
	PptxSmartArtConnection,
	PptxSmartArtNode,
	SmartArtNodeCustomLayout,
	XmlObject,
} from '../types';

type NullableAttribute = string | null | undefined;

/** Reads `@_name` attributes of a `fast-xml-parser` node for the neutral diagram parsers. */
const objectAttributeReader = (node: XmlObject) => (name: string) => {
	const value = node['@_' + name];
	return value === undefined ? undefined : String(value ?? '');
};

export interface SmartArtDataModelIssue {
	code:
		| 'POINT_ID_REQUIRED'
		| 'POINT_ID_DUPLICATE'
		| 'CONNECTION_ATTRIBUTE_REQUIRED'
		| 'CONNECTION_ID_DUPLICATE'
		| 'CONNECTION_ENDPOINT_MISSING';
	message: string;
}

function optionalString(value: unknown): string | undefined {
	const text = String(value ?? '').trim();
	return text.length > 0 ? text : undefined;
}

function applyNullableAttribute(xml: XmlObject, key: string, value: NullableAttribute): void {
	if (value === undefined) {
		return;
	}
	if (value === null || value.trim().length === 0) {
		delete xml[key];
		return;
	}
	xml[key] = value;
}

/** Parse the typed CT_Cxn attributes while leaving its XML object untouched. */
export function parseSmartArtConnection(connection: XmlObject): PptxSmartArtConnection | undefined {
	return parseConnectionAttributes(objectAttributeReader(connection));
}

/** Apply editable CT_Pt attributes without disturbing unknown attributes/children. */
export function applySmartArtPointAttributes(xml: XmlObject, node: PptxSmartArtNode): void {
	applyNullableAttribute(xml, '@_cxnId', node.connectionId);
}

/** Apply editable CT_Cxn identifiers without disturbing unknown XML content. */
export function applySmartArtConnectionAttributes(
	xml: XmlObject,
	connection: PptxSmartArtConnection,
	fallbackModelId: () => string,
): void {
	applyNullableAttribute(xml, '@_modelId', connection.modelId);
	if (!xml['@_modelId']) {
		xml['@_modelId'] = fallbackModelId();
	}
	xml['@_srcId'] = connection.sourceId;
	xml['@_destId'] = connection.destId;
	applyNullableAttribute(xml, '@_parTransId', connection.parentTransitionId);
	applyNullableAttribute(xml, '@_sibTransId', connection.siblingTransitionId);
	applyNullableAttribute(xml, '@_presId', connection.presentationId);
}

/**
 * Parse the manual layout override attributes (`cust*`) from a `dgm:pt`'s
 * `dgm:prSet` element, or `undefined` when `prSet` is absent or carries none of
 * them. These capture a drag/resize/rotate/flip the user performed on a
 * `type="pres"` presentation point in PowerPoint's own diagram editor.
 */
export function parseSmartArtPointCustomLayout(
	prSet: XmlObject | undefined,
): SmartArtNodeCustomLayout | undefined {
	return prSet ? parseCustomLayoutAttributes(objectAttributeReader(prSet)) : undefined;
}

function childrenByLocalName(parent: XmlObject | undefined, name: string): XmlObject[] {
	const key = Object.keys(parent ?? {}).find((entry) => entry.split(':').pop() === name);
	const value = key ? parent?.[key] : undefined;
	if (!value) {
		return [];
	}
	return (Array.isArray(value) ? value : [value]).filter(
		(entry): entry is XmlObject => typeof entry === 'object' && entry !== null,
	);
}

/** Validate the required CT_Pt/CT_Cxn identifiers and graph references. */
export function validateSmartArtDataModelCore(dataModel: XmlObject): SmartArtDataModelIssue[] {
	const issues: SmartArtDataModelIssue[] = [];
	const pointList = childrenByLocalName(dataModel, 'ptLst')[0];
	const connectionList = childrenByLocalName(dataModel, 'cxnLst')[0];
	const pointIds = new Set<string>();
	for (const point of childrenByLocalName(pointList, 'pt')) {
		const id = optionalString(point['@_modelId']);
		if (!id) {
			issues.push({ code: 'POINT_ID_REQUIRED', message: 'dgm:pt requires modelId.' });
		} else if (pointIds.has(id)) {
			issues.push({ code: 'POINT_ID_DUPLICATE', message: `Duplicate dgm:pt modelId: ${id}.` });
		} else {
			pointIds.add(id);
		}
	}

	const connectionIds = new Set<string>();
	for (const connection of childrenByLocalName(connectionList, 'cxn')) {
		const id = optionalString(connection['@_modelId']);
		const sourceId = optionalString(connection['@_srcId']);
		const destId = optionalString(connection['@_destId']);
		for (const [attribute, value] of [
			['modelId', id],
			['srcId', sourceId],
			['destId', destId],
		]) {
			if (!value) {
				issues.push({
					code: 'CONNECTION_ATTRIBUTE_REQUIRED',
					message: `dgm:cxn requires ${attribute}.`,
				});
			}
		}
		if (id && connectionIds.has(id)) {
			issues.push({
				code: 'CONNECTION_ID_DUPLICATE',
				message: `Duplicate dgm:cxn modelId: ${id}.`,
			});
		} else if (id) {
			connectionIds.add(id);
		}
		for (const [attribute, value] of [
			['srcId', sourceId],
			['destId', destId],
		]) {
			if (value && !pointIds.has(value)) {
				issues.push({
					code: 'CONNECTION_ENDPOINT_MISSING',
					message: `dgm:cxn ${attribute} references missing point: ${value}.`,
				});
			}
		}
	}
	return issues;
}
