/**
 * XML reference transform processing for digital signature verification.
 *
 * Node-only - depends on `@xmldom/xmldom` for DOM parsing and
 * `xml-crypto` (via xml-canonicalization) for C14N transforms.
 */

import {
	OPC_DIGITAL_SIGNATURE_NS,
	OPC_RELATIONSHIPS_NS,
	OPC_RELATIONSHIP_TRANSFORM,
	SUPPORTED_XML_CANON_TRANSFORMS,
	XMLDSIG_NS,
} from '../core/utils/signature-constants';
import type {
	PackageRelationship,
	ParsedReferenceTransform,
	ReferenceTransformResult,
} from '../core/utils/signature-types';
import { escapeXmlAttr } from '../core/utils/signature-xml-utils';
import {
	canonicalizeNode,
	childElements,
	getNodeLocalName,
	parseStrictXml,
} from './xml-canonicalization';
import type { XmlElement } from './xml-canonicalization';

/**
 * Parse `<ds:Transform>` elements from a `<ds:Reference>` node.
 */
export function extractReferenceTransforms(referenceNode: XmlElement): ParsedReferenceTransform[] {
	const transforms: ParsedReferenceTransform[] = [];
	const transformNodes = referenceNode.getElementsByTagNameNS(XMLDSIG_NS, 'Transform');
	for (let index = 0; index < transformNodes.length; index += 1) {
		const transformNode = transformNodes.item(index);
		const algorithm = transformNode?.getAttribute('Algorithm')?.trim();
		if (!transformNode || !algorithm) {
			continue;
		}
		const parameters = childElements(transformNode).filter(
			(element) => element.namespaceURI === OPC_DIGITAL_SIGNATURE_NS,
		);
		const parameterValues = (localName: string, attribute: string): string[] =>
			parameters
				.filter((element) => getNodeLocalName(element) === localName)
				.map((element) => element.getAttribute(attribute)?.trim() ?? '')
				.filter((value) => value.length > 0);
		transforms.push({
			algorithm,
			relationshipReferenceIds: parameterValues('RelationshipReference', 'SourceId'),
			relationshipGroupTypes: parameterValues('RelationshipsGroupReference', 'SourceType'),
		});
	}
	return transforms;
}

/** Read the `Relationship` elements of a relationships part, in document order. */
export function parseRelationshipsPart(xmlText: string): PackageRelationship[] {
	const root = parseStrictXml(xmlText).documentElement!;
	if (root.namespaceURI !== OPC_RELATIONSHIPS_NS || getNodeLocalName(root) !== 'Relationships') {
		throw new Error('Not a relationships part.');
	}
	return childElements(root)
		.filter(
			(element) =>
				element.namespaceURI === OPC_RELATIONSHIPS_NS &&
				getNodeLocalName(element) === 'Relationship',
		)
		.map((element) => ({
			id: element.getAttribute('Id') ?? '',
			type: element.getAttribute('Type') ?? '',
			target: element.getAttribute('Target') ?? '',
			mode: element.getAttribute('TargetMode') === 'External' ? 'External' : 'Internal',
		}));
}

const compareCodePoints = (left: string, right: string): number =>
	left < right ? -1 : left > right ? 1 : 0;

/**
 * The relationships transform of ECMA-376 Part 2: keep the relationships the signer selected by
 * `SourceId` or `SourceType`, sort them by `Id`, state `TargetMode` explicitly and drop everything
 * else from the part. The result is rebuilt from the parsed values, so whitespace and foreign markup
 * in the part cannot change it.
 */
function applyRelationshipTransform(
	xmlText: string,
	transform: ParsedReferenceTransform,
): { xml: string; selected: PackageRelationship[] } {
	const ids = new Set(transform.relationshipReferenceIds);
	const types = new Set(transform.relationshipGroupTypes);
	const selected = parseRelationshipsPart(xmlText)
		.filter((relationship) => ids.has(relationship.id) || types.has(relationship.type))
		.sort((left, right) => compareCodePoints(left.id, right.id));
	const body = selected
		.map(
			(relationship) =>
				`<Relationship Id="${escapeXmlAttr(relationship.id)}"` +
				` Target="${escapeXmlAttr(relationship.target)}"` +
				` TargetMode="${relationship.mode}"` +
				` Type="${escapeXmlAttr(relationship.type)}"></Relationship>`,
		)
		.join('');
	return {
		xml: `<Relationships xmlns="${OPC_RELATIONSHIPS_NS}">${body}</Relationships>`,
		selected,
	};
}

/**
 * Apply a chain of transforms to binary part data.
 * Supports the OPC relationships transform and XML canonicalization algorithms.
 */
export function applyReferenceTransforms(
	partBytes: Uint8Array,
	transforms: ParsedReferenceTransform[],
): ReferenceTransformResult {
	let transformedBytes = partBytes;
	let selectedRelationships: PackageRelationship[] | undefined;
	const unsupportedAlgorithms: string[] = [];
	for (const transform of transforms) {
		const text = Buffer.from(transformedBytes).toString('utf8');
		try {
			let output: string;
			if (transform.algorithm === OPC_RELATIONSHIP_TRANSFORM) {
				const result = applyRelationshipTransform(text, transform);
				output = result.xml;
				selectedRelationships = result.selected;
			} else if (SUPPORTED_XML_CANON_TRANSFORMS.has(transform.algorithm)) {
				output = canonicalizeNode(parseStrictXml(text), transform.algorithm);
			} else {
				unsupportedAlgorithms.push(transform.algorithm);
				continue;
			}
			transformedBytes = new Uint8Array(Buffer.from(output, 'utf8'));
		} catch {
			unsupportedAlgorithms.push(transform.algorithm);
		}
	}
	return { data: transformedBytes, unsupportedAlgorithms, selectedRelationships };
}
