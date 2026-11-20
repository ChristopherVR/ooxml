/**
 * The reference transforms of an OPC signature: the ECMA-376 Part 2 relationships transform and
 * XML canonicalization. Platform-neutral.
 */

import { decodeXmlBytes } from '../../../xml/index';
import { asciiLowerCase } from '../../part-names';
import { parseRelationshipsPart } from '../../relationships';
import {
	OPC_DIGITAL_SIGNATURE_NS,
	OPC_RELATIONSHIPS_NS,
	OPC_RELATIONSHIP_TRANSFORM,
	SUPPORTED_XML_CANON_TRANSFORMS,
	XMLDSIG_NS,
} from '../constants';
import type {
	PackageRelationship,
	ParsedReferenceTransform,
	ReferenceTransformResult,
} from '../types';
import { escapeXmlAttr } from '../xml-utils';
import { canonicalize } from './c14n';
import { childElements, localNameOf, parseStrictXml } from './dom';

/**
 * Parse `<ds:Transform>` elements from a `<ds:Reference>` node.
 */
export function extractReferenceTransforms(referenceNode: Element): ParsedReferenceTransform[] {
	const transforms: ParsedReferenceTransform[] = [];
	const transformNodes = referenceNode.getElementsByTagNameNS(XMLDSIG_NS, 'Transform');
	for (let index = 0; index < transformNodes.length; index += 1) {
		const transformNode = transformNodes.item(index);
		const algorithm = transformNode?.getAttribute('Algorithm')?.trim();
		if (!transformNode || !algorithm) {
			continue;
		}
		const children = childElements(transformNode);
		const parameters = children.filter(
			(element) => element.namespaceURI === OPC_DIGITAL_SIGNATURE_NS,
		);
		const parameterValues = (localName: string, attribute: string): string[] =>
			parameters
				.filter((element) => localNameOf(element) === localName)
				.map((element) => element.getAttribute(attribute)?.trim() ?? '')
				.filter((value) => value.length > 0);
		transforms.push({
			algorithm,
			relationshipReferenceIds: parameterValues('RelationshipReference', 'SourceId'),
			relationshipGroupTypes: parameterValues('RelationshipsGroupReference', 'SourceType'),
			// Other parameters (an exclusive canonicalization prefix list, say) change the output,
			// and this verifier does not apply them.
			unsupportedParameters: parameters.length < children.length,
		});
	}
	return transforms;
}

const compareCodePoints = (left: string, right: string): number =>
	left < right ? -1 : left > right ? 1 : 0;

/**
 * The relationships transform of ECMA-376 Part 2: keep the relationships the signer selected by
 * `SourceId` or `SourceType` (matched ASCII case-insensitively), sort them by `Id`, state
 * `TargetMode` explicitly and drop everything else. The result is rebuilt from the strictly parsed
 * values, so whitespace cannot change it and a part the strict reader rejects has no output.
 */
function applyRelationshipTransform(
	xmlText: string,
	transform: ParsedReferenceTransform,
): { xml: string; selected: PackageRelationship[] } {
	const ids = new Set(transform.relationshipReferenceIds.map(asciiLowerCase));
	const types = new Set(transform.relationshipGroupTypes.map(asciiLowerCase));
	const selected = parseRelationshipsPart(xmlText)
		.filter(
			(relationship) =>
				ids.has(asciiLowerCase(relationship.id)) || types.has(asciiLowerCase(relationship.type)),
		)
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
		try {
			let output: string;
			if (transform.unsupportedParameters) {
				unsupportedAlgorithms.push(transform.algorithm);
				continue;
			} else if (transform.algorithm === OPC_RELATIONSHIP_TRANSFORM) {
				const result = applyRelationshipTransform(decodeXmlBytes(transformedBytes), transform);
				output = result.xml;
				selectedRelationships = result.selected;
			} else if (SUPPORTED_XML_CANON_TRANSFORMS.has(transform.algorithm)) {
				output = canonicalize(
					parseStrictXml(decodeXmlBytes(transformedBytes)),
					transform.algorithm,
				);
			} else {
				unsupportedAlgorithms.push(transform.algorithm);
				continue;
			}
			transformedBytes = new TextEncoder().encode(output);
		} catch {
			unsupportedAlgorithms.push(transform.algorithm);
		}
	}
	return { data: transformedBytes, unsupportedAlgorithms, selectedRelationships };
}
