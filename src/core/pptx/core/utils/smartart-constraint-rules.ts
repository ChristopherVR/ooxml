import {
	constraintAttributes,
	parseConstraintAttributes,
	parseRuleAttributes,
	ruleAttributes,
	validateSmartArtConstraintRules,
} from '../../../diagram/layout/smartart-constraint-rules';
import type { DiagramAttributeValue } from '../../../diagram/layout/smartart-constraint-rules';
import type {
	PptxSmartArtConstraint,
	PptxSmartArtLayoutNode,
	PptxSmartArtNumericRule,
	XmlObject,
} from '../types';
import { cloneXmlObject } from './clone-utils';

export { validateSmartArtConstraintRules } from '../../../diagram/layout/smartart-constraint-rules';

type LocalName = (key: string) => string;

function keyOf(node: XmlObject, name: string, localName: LocalName): string | undefined {
	return Object.keys(node).find((key) => localName(key) === name);
}

function child(node: XmlObject, name: string, localName: LocalName): XmlObject | undefined {
	const key = keyOf(node, name, localName);
	const value = key ? node[key] : undefined;
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as XmlObject)
		: undefined;
}

function children(node: XmlObject, name: string, localName: LocalName): XmlObject[] {
	const key = keyOf(node, name, localName);
	const value = key ? node[key] : undefined;
	if (Array.isArray(value)) {
		return value as XmlObject[];
	}
	return value && typeof value === 'object' ? [value as XmlObject] : [];
}

function attr(node: XmlObject, name: string, localName: LocalName): string | undefined {
	const key = Object.keys(node).find(
		(candidate) => candidate.startsWith('@_') && localName(candidate.replace(/^@_/u, '')) === name,
	);
	return key && node[key] !== undefined ? String(node[key]) : undefined;
}

/** Reads attributes by local name for the neutral `diagram` parsers. */
function reader(node: XmlObject, localName: LocalName): (name: string) => string | undefined {
	return (name) => attr(node, name, localName);
}

export function parseConstraint(node: XmlObject, localName: LocalName): PptxSmartArtConstraint {
	return parseConstraintAttributes(reader(node, localName), cloneXmlObject(node));
}

export function parseRule(node: XmlObject, localName: LocalName): PptxSmartArtNumericRule {
	return parseRuleAttributes(reader(node, localName), cloneXmlObject(node));
}

export function parseSmartArtConstraintRules(
	node: XmlObject,
	localName: LocalName,
): Pick<PptxSmartArtLayoutNode, 'constraints' | 'rules'> {
	const constraintList = child(node, 'constrLst', localName);
	const ruleList = child(node, 'ruleLst', localName);
	return {
		constraints: constraintList
			? children(constraintList, 'constr', localName).map((item) =>
					parseConstraint(item, localName),
				)
			: undefined,
		rules: ruleList
			? children(ruleList, 'rule', localName).map((item) => parseRule(item, localName))
			: undefined,
	};
}

function setAttributes(
	node: XmlObject,
	values: readonly DiagramAttributeValue[],
	localName: LocalName,
): void {
	for (const [name, value] of values) {
		const key =
			Object.keys(node).find(
				(candidate) =>
					candidate.startsWith('@_') && localName(candidate.replace(/^@_/u, '')) === name,
			) ?? `@_${name}`;
		if (value === undefined) {
			delete node[key];
		} else {
			node[key] = value;
		}
	}
}

function itemPrefix(listKey: string | undefined): string {
	return listKey?.includes(':') ? `${listKey.slice(0, listKey.indexOf(':'))}:` : 'dgm:';
}

/** Merge typed constraints and rules without dropping foreign attributes or extension content. */
export function applySmartArtConstraintRules(
	node: XmlObject,
	value: PptxSmartArtLayoutNode,
	localName: LocalName,
): boolean {
	if (validateSmartArtConstraintRules(value).length > 0) {
		return false;
	}
	for (const [listName, itemName, values] of [
		['constrLst', 'constr', value.constraints],
		['ruleLst', 'rule', value.rules],
	] as const) {
		if (values === undefined) {
			continue;
		}
		const listKey = keyOf(node, listName, localName) ?? `dgm:${listName}`;
		const list = child(node, listName, localName) ?? {};
		const oldItems = children(list, itemName, localName);
		const itemKey = keyOf(list, itemName, localName) ?? `${itemPrefix(listKey)}${itemName}`;
		for (const key of Object.keys(list)) {
			if (localName(key) === itemName) {
				delete list[key];
			}
		}
		list[itemKey] = values.map((item, index) => {
			const target = cloneXmlObject(oldItems[index]) ?? cloneXmlObject(item.rawXml) ?? {};
			setAttributes(
				target,
				itemName === 'constr'
					? constraintAttributes(item as PptxSmartArtConstraint)
					: ruleAttributes(item as PptxSmartArtNumericRule),
				localName,
			);
			return target;
		});
		node[listKey] = list;
	}
	return true;
}
