import type {
	PptxSmartArtChoose,
	PptxSmartArtForEach,
	PptxSmartArtIteratorAttributes,
	PptxSmartArtLayoutNode,
	PptxSmartArtWhen,
	XmlObject,
} from '../types';
import {
	parseIteratorAttributes,
	parseWhenAttributes,
} from '../../../diagram/layout/smartart-layout-control-flow';

type LocalName = (key: string) => string;

function children(node: XmlObject, name: string, localName: LocalName): XmlObject[] {
	const key = Object.keys(node).find((candidate) => localName(candidate) === name);
	const value = key ? node[key] : undefined;
	return Array.isArray(value) ? value : value && typeof value === 'object' ? [value] : [];
}

/** Reads a `fast-xml-parser` node's attributes for the neutral `diagram` parsers. */
function attributes(node: XmlObject): (name: string) => string | undefined {
	return (name) => {
		const value = node[`@_${name}`];
		return value === undefined || value === null ? undefined : String(value);
	};
}

function optionalString(value: unknown): string | undefined {
	const result = String(value ?? '').trim();
	return result.length > 0 ? result : undefined;
}

/**
 * Parse CT_Iterate's shared attribute set (`axis`/`ptType`/`st`/`cnt`/`step`/
 * `hideLastTrans`). Exposed for `smartart-layout-definition.ts`, which reuses
 * it verbatim for `dgm:presOf` (CT_PresentationOf extends CT_Iterate).
 */
export function parseIterator(node: XmlObject): PptxSmartArtIteratorAttributes {
	return parseIteratorAttributes(attributes(node));
}

/** A `dgm:if` of the object tree (see `parseWhenAttributes` in `diagram`). */
export function parseWhen(node: XmlObject): PptxSmartArtWhen | undefined {
	return parseWhenAttributes(attributes(node), node);
}

export function parseSmartArtControlFlow(
	node: XmlObject,
	localName: LocalName,
): Pick<PptxSmartArtLayoutNode, 'forEach' | 'choose'> {
	const forEach = children(node, 'forEach', localName).map((entry): PptxSmartArtForEach => ({
		...parseIterator(entry),
		rawXml: entry,
	}));
	const choose = children(node, 'choose', localName).map((entry): PptxSmartArtChoose => {
		const otherwiseNode = children(entry, 'else', localName)[0];
		return {
			name: optionalString(entry['@_name']),
			when: children(entry, 'if', localName)
				.map(parseWhen)
				.filter((value) => value !== undefined),
			otherwise: otherwiseNode
				? { name: optionalString(otherwiseNode['@_name']), rawXml: otherwiseNode }
				: undefined,
			rawXml: entry,
		};
	});
	return {
		...(forEach.length ? { forEach } : {}),
		...(choose.length ? { choose } : {}),
	};
}

function set(node: XmlObject, name: string, value: string | undefined): void {
	if (value === undefined) {
		delete node[`@_${name}`];
	} else {
		node[`@_${name}`] = value;
	}
}

function applyIterator(node: XmlObject, value: PptxSmartArtIteratorAttributes): void {
	set(node, 'name', value.name);
	set(node, 'ref', value.reference);
	set(node, 'axis', value.axis?.join(' '));
	set(node, 'ptType', value.pointTypes?.join(' '));
	set(
		node,
		'hideLastTrans',
		value.hideLastTransition?.map((entry) => (entry ? '1' : '0')).join(' '),
	);
	set(node, 'st', value.start?.join(' '));
	set(node, 'cnt', value.count?.join(' '));
	set(node, 'step', value.step?.join(' '));
}

function applyWhen(value: PptxSmartArtWhen): XmlObject {
	const node = { ...(value.rawXml ?? {}) };
	applyIterator(node, value);
	set(node, 'func', value.function);
	set(node, 'arg', value.argument);
	set(node, 'op', value.operator);
	set(node, 'val', value.value);
	return node;
}

function applyChoose(value: PptxSmartArtChoose, localName: LocalName): XmlObject {
	const node = { ...(value.rawXml ?? {}) };
	set(node, 'name', value.name);
	const ifKey = Object.keys(node).find((key) => localName(key) === 'if') ?? 'dgm:if';
	node[ifKey] = value.when.map(applyWhen);
	const elseKey = Object.keys(node).find((key) => localName(key) === 'else') ?? 'dgm:else';
	if (value.otherwise === null) {
		delete node[elseKey];
	} else if (value.otherwise) {
		const otherwise = { ...(value.otherwise.rawXml ?? {}) };
		set(otherwise, 'name', value.otherwise.name);
		node[elseKey] = otherwise;
	}
	return node;
}

function replaceChildren(
	node: XmlObject,
	name: string,
	values: XmlObject[] | undefined,
	localName: LocalName,
): void {
	if (values === undefined) {
		return;
	}
	const key = Object.keys(node).find((candidate) => localName(candidate) === name) ?? `dgm:${name}`;
	if (values.length === 0) {
		delete node[key];
	} else {
		node[key] = values;
	}
}

export function applySmartArtControlFlow(
	node: XmlObject,
	value: PptxSmartArtLayoutNode,
	localName: LocalName,
): void {
	replaceChildren(
		node,
		'forEach',
		value.forEach?.map((entry) => {
			const target = { ...(entry.rawXml ?? {}) };
			applyIterator(target, entry);
			return target;
		}),
		localName,
	);
	replaceChildren(
		node,
		'choose',
		value.choose?.map((entry) => applyChoose(entry, localName)),
		localName,
	);
}

export { validateSmartArtControlFlow } from '../../../diagram/layout/smartart-layout-control-flow';
