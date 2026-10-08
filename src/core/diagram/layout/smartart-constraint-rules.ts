/**
 * DiagramML `dgm:constr` (CT_Constraint) and `dgm:rule` (CT_NumericRule)
 * attributes read through an {@link AttributeReader}, their validation, and
 * the attribute values a writer puts back. Extracted from
 * `pptx/core/utils/smartart-constraint-rules.ts`; pptx keeps the object-tree
 * list lookup and the writer that merges these values into its parts.
 */

import type { AttributeReader } from '../../drawingml/types';
import type {
	DiagramConstraint,
	DiagramConstraintOperator,
	DiagramConstraintPointType,
	DiagramConstraintRelationship,
	DiagramConstraintTarget,
	DiagramLayoutNode,
	DiagramNumericRule,
} from '../model';
import {
	SMART_ART_CONSTRAINT_OPERATORS,
	SMART_ART_CONSTRAINT_TYPES,
	SMART_ART_POINT_TYPES,
	SMART_ART_RELATIONSHIPS,
} from './smartart-constraint-values';

/** An `xsd:double` attribute (`NaN`, `INF` and `-INF` included), `undefined` when absent or invalid. */
export function parseXsdDouble(value: string | undefined): number | undefined {
	if (value === undefined || value.trim() === '') {
		return undefined;
	}
	if (value === 'NaN') {
		return Number.NaN;
	}
	if (value === 'INF') {
		return Number.POSITIVE_INFINITY;
	}
	if (value === '-INF') {
		return Number.NEGATIVE_INFINITY;
	}
	const parsed = Number(value);
	return Number.isNaN(parsed) ? undefined : parsed;
}

/** The `xsd:double` lexical form of `value`. */
export function formatXsdDouble(value: number): string {
	if (Number.isNaN(value)) {
		return 'NaN';
	}
	if (value === Number.POSITIVE_INFINITY) {
		return 'INF';
	}
	if (value === Number.NEGATIVE_INFINITY) {
		return '-INF';
	}
	return String(value);
}

function parseTarget(read: AttributeReader): DiagramConstraintTarget {
	return {
		for: read('for') as DiagramConstraintRelationship | undefined,
		forName: read('forName'),
		pointType: read('ptType') as DiagramConstraintPointType | undefined,
	};
}

/** A `dgm:constr`, keeping `rawXml` as the reader's round-trip slot. */
export function parseConstraintAttributes<R>(
	read: AttributeReader,
	rawXml: R | undefined,
): DiagramConstraint<R> {
	return {
		type: read('type') ?? '',
		...parseTarget(read),
		referenceType: read('refType'),
		referenceFor: read('refFor') as DiagramConstraintRelationship | undefined,
		referenceForName: read('refForName'),
		referencePointType: read('refPtType') as DiagramConstraintPointType | undefined,
		operator: read('op') as DiagramConstraintOperator | undefined,
		value: parseXsdDouble(read('val')),
		factor: parseXsdDouble(read('fact')),
		rawXml,
	};
}

/** A `dgm:rule`, keeping `rawXml` as the reader's round-trip slot. */
export function parseRuleAttributes<R>(
	read: AttributeReader,
	rawXml: R | undefined,
): DiagramNumericRule<R> {
	return {
		type: read('type') ?? '',
		...parseTarget(read),
		value: parseXsdDouble(read('val')),
		factor: parseXsdDouble(read('fact')),
		max: parseXsdDouble(read('max')),
		rawXml,
	};
}

function validateTarget(value: DiagramConstraint | DiagramNumericRule, path: string): string[] {
	const errors: string[] = [];
	if (!SMART_ART_CONSTRAINT_TYPES.has(value.type)) {
		errors.push(`${path}.type is invalid`);
	}
	if (value.for !== undefined && !SMART_ART_RELATIONSHIPS.has(value.for)) {
		errors.push(`${path}.for is invalid`);
	}
	if (value.pointType !== undefined && !SMART_ART_POINT_TYPES.has(value.pointType)) {
		errors.push(`${path}.pointType is invalid`);
	}
	return errors;
}

/** Enumeration errors in a layout node's constraints and rules. */
export function validateSmartArtConstraintRules(value: DiagramLayoutNode): string[] {
	const errors: string[] = [];
	value.constraints?.forEach((item, index) => {
		const path = `constraints[${index}]`;
		errors.push(...validateTarget(item, path));
		if (item.referenceType !== undefined && !SMART_ART_CONSTRAINT_TYPES.has(item.referenceType)) {
			errors.push(`${path}.referenceType is invalid`);
		}
		if (item.referenceFor !== undefined && !SMART_ART_RELATIONSHIPS.has(item.referenceFor)) {
			errors.push(`${path}.referenceFor is invalid`);
		}
		if (
			item.referencePointType !== undefined &&
			!SMART_ART_POINT_TYPES.has(item.referencePointType)
		) {
			errors.push(`${path}.referencePointType is invalid`);
		}
		if (item.operator !== undefined && !SMART_ART_CONSTRAINT_OPERATORS.has(item.operator)) {
			errors.push(`${path}.operator is invalid`);
		}
	});
	value.rules?.forEach((item, index) => errors.push(...validateTarget(item, `rules[${index}]`)));
	return errors;
}

/** One attribute a writer sets (a string) or removes (`undefined`), in write order. */
export type DiagramAttributeValue = readonly [name: string, value: string | undefined];

function text(value: string | number | undefined): string | undefined {
	return typeof value === 'number' ? formatXsdDouble(value) : value;
}

function targetAttributes(value: DiagramConstraint | DiagramNumericRule): DiagramAttributeValue[] {
	return [
		['type', value.type],
		['for', value.for],
		['forName', value.forName],
		['ptType', value.pointType],
		['val', text(value.value)],
		['fact', text(value.factor)],
	];
}

/** The `dgm:constr` attributes a writer sets for `value`, in write order. */
export function constraintAttributes(value: DiagramConstraint): DiagramAttributeValue[] {
	return [
		...targetAttributes(value),
		['refType', value.referenceType],
		['refFor', value.referenceFor],
		['refForName', value.referenceForName],
		['refPtType', value.referencePointType],
		['op', value.operator],
	];
}

/** The `dgm:rule` attributes a writer sets for `value`, in write order. */
export function ruleAttributes(value: DiagramNumericRule): DiagramAttributeValue[] {
	return [...targetAttributes(value), ['max', text(value.max)]];
}
