/**
 * DiagramML control-flow attributes (`dgm:forEach`, `dgm:if`, `dgm:presOf`,
 * all CT_Iterate based) read through an {@link AttributeReader}, so the same
 * parser serves the pptx object tree and the ordered-XML walkers, plus the
 * format-neutral validation of the typed model. Extracted from
 * `pptx/core/utils/smartart-layout-control-flow.ts` (pptx keeps the object
 * tree child lookup and the writers).
 */

import type { AttributeReader } from '../../drawingml/types';
import type { DiagramIteratorAttributes, DiagramLayoutNode, DiagramWhen } from '../model';
import type { OrderedXmlElement } from './smartart-choose-xml';

const UINT_MAX = 4_294_967_295;

function optionalString(value: string | undefined): string | undefined {
	const result = (value ?? '').trim();
	return result.length > 0 ? result : undefined;
}

function strings(value: string | undefined): string[] | undefined {
	const values = optionalString(value)?.split(/\s+/u);
	return values?.length ? values : undefined;
}

function booleans(value: string | undefined): boolean[] | undefined {
	const values = strings(value);
	if (!values || values.some((entry) => !['0', '1', 'true', 'false'].includes(entry))) {
		return undefined;
	}
	return values.map((entry) => entry === '1' || entry === 'true');
}

function integers(value: string | undefined, unsigned = false): number[] | undefined {
	const values = strings(value);
	if (!values) {
		return undefined;
	}
	const parsed = values.map(Number);
	if (
		parsed.some(
			(entry) =>
				!Number.isInteger(entry) ||
				entry < (unsigned ? 0 : -2_147_483_648) ||
				entry > (unsigned ? UINT_MAX : 2_147_483_647),
		)
	) {
		return undefined;
	}
	return parsed;
}

/** Reads an ordered element's attributes by local name. */
export function orderedAttributes(element: OrderedXmlElement): AttributeReader {
	return (name) => element.attrs[name];
}

/**
 * CT_Iterate's shared attribute set (`name`/`ref`/`axis`/`ptType`/
 * `hideLastTrans`/`st`/`cnt`/`step`), also used verbatim for `dgm:presOf`.
 */
export function parseIteratorAttributes(read: AttributeReader): DiagramIteratorAttributes {
	return {
		name: optionalString(read('name')),
		reference: optionalString(read('ref')),
		axis: strings(read('axis')),
		pointTypes: strings(read('ptType')),
		hideLastTransition: booleans(read('hideLastTrans')),
		start: integers(read('st')),
		count: integers(read('cnt'), true),
		step: integers(read('step')),
	};
}

/** A `dgm:if` (CT_When), or `undefined` when `func`, `op` or `val` is missing. */
export function parseWhenAttributes<R>(
	read: AttributeReader,
	rawXml: R,
): DiagramWhen<R> | undefined {
	const func = optionalString(read('func'));
	const operator = optionalString(read('op'));
	const value = optionalString(read('val'));
	return func && operator && value
		? {
				...parseIteratorAttributes(read),
				function: func,
				argument: optionalString(read('arg')),
				operator,
				value,
				rawXml,
			}
		: undefined;
}

/** A `dgm:if` element of the ordered tree, its raw slot the element itself. */
export function parseOrderedWhen(
	element: OrderedXmlElement,
): DiagramWhen<OrderedXmlElement> | undefined {
	return parseWhenAttributes(orderedAttributes(element), element);
}

/** Range and required-field errors in a layout node's `forEach`/`choose` model. */
export function validateSmartArtControlFlow(node: DiagramLayoutNode): string[] {
	const errors: string[] = [];
	const validateIterator = (value: DiagramIteratorAttributes, path: string): void => {
		for (const field of ['start', 'step'] as const) {
			if (
				value[field]?.some(
					(entry) => !Number.isInteger(entry) || entry < -2_147_483_648 || entry > 2_147_483_647,
				)
			) {
				errors.push(`${path}.${field} values must be signed 32-bit integers`);
			}
		}
		if (value.count?.some((entry) => !Number.isInteger(entry) || entry < 0 || entry > UINT_MAX)) {
			errors.push(`${path}.count values must be unsigned 32-bit integers`);
		}
	};
	node.forEach?.forEach((value, index) => validateIterator(value, `forEach[${index}]`));
	node.choose?.forEach((choose, chooseIndex) => {
		if (choose.when.length === 0) {
			errors.push(`choose[${chooseIndex}].when requires at least one branch`);
		}
		choose.when.forEach((branch, branchIndex) => {
			const path = `choose[${chooseIndex}].when[${branchIndex}]`;
			validateIterator(branch, path);
			for (const field of ['function', 'operator', 'value'] as const) {
				if (!branch[field].trim()) {
					errors.push(`${path}.${field} is required`);
				}
			}
		});
	});
	return errors;
}
