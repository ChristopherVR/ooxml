// Shared helpers of the `cx:chartSpace` DOM parser: chartex-namespace lookups, typed attribute
// readers that report malformed values instead of throwing, and the not-modelled child report.
// Shape and text properties go through the classic chart readers (`../parse-util`), which are
// namespace-neutral below the wrapper element.
import { NS, buildXml, elements, first, type XmlElement } from '../../xml/index';
import type {
	ChartLines,
	ChartNumberFormat,
	ChartShapeProperties,
	ChartTextBody,
} from '../model-series';
import { attribute, chartTextBody, readShapeFields, type ChartParseContext } from '../parse-util';
import { innerXml } from '../xml-fragment';
import type { ChartExAutoNumber, ChartExFormula } from './model';

export { attribute, assignDefined, type ChartParseContext } from '../parse-util';

/** The first chartex-namespace child (lenient: un-namespaced test XML matches too). */
export const cxChild = (parent: ParentNode | null | undefined, local: string) =>
	first(parent, local, NS.cx);

/** Every chartex-namespace child named `local`. */
export const cxChildren = (parent: ParentNode | null | undefined, local: string) =>
	parent ? elements(parent).filter((el) => el.localName === local && isCx(el)) : [];

const isCx = (element: XmlElement) => !element.namespaceURI || element.namespaceURI === NS.cx;

const path = (element: XmlElement, name?: string) =>
	`cx:${element.localName}${name ? `/@${name}` : ''}`;

function malformed(context: ChartParseContext, where: string, value: string): void {
	context.issues.push({
		code: 'CHART_VALUE_INVALID',
		message: `${where} has an unreadable value "${value}".`,
	});
}

/** An `xsd:boolean` attribute; malformed values are reported and dropped. */
export function boolAttribute(
	context: ChartParseContext,
	element: XmlElement | undefined,
	name: string,
): boolean | undefined {
	const raw = attribute(element, name);
	if (!element || raw === undefined) return undefined;
	const value = raw.trim().toLowerCase();
	if (value === '1' || value === 'true') return true;
	if (value === '0' || value === 'false') return false;
	malformed(context, path(element, name), raw);
	return undefined;
}

/** A numeric attribute; malformed values are reported and dropped. */
export function numberAttribute(
	context: ChartParseContext,
	element: XmlElement | undefined,
	name: string,
): number | undefined {
	const raw = attribute(element, name);
	if (!element || raw === undefined) return undefined;
	const value = raw.trim() ? Number(raw) : Number.NaN;
	if (Number.isFinite(value)) return value;
	malformed(context, path(element, name), raw);
	return undefined;
}

/** A non-negative integer attribute (`@idx`, `@id`); other values are reported and dropped. */
export function indexAttribute(
	context: ChartParseContext,
	element: XmlElement | undefined,
	name: string,
): number | undefined {
	const value = numberAttribute(context, element, name);
	if (value === undefined || (Number.isInteger(value) && value >= 0)) return value;
	malformed(context, path(element!, name), attribute(element, name) ?? '');
	return undefined;
}

/** A double-or-`auto` attribute (`@gapWidth`, `@min`, `@underflow`...). */
export function autoNumberAttribute(
	context: ChartParseContext,
	element: XmlElement | undefined,
	name: string,
): ChartExAutoNumber | undefined {
	return attribute(element, name)?.trim() === 'auto'
		? 'auto'
		: numberAttribute(context, element, name);
}

/** `@val` of the chartex child `local`, as a number. */
export const valNumber = (
	context: ChartParseContext,
	parent: XmlElement | undefined,
	local: string,
): number | undefined => numberAttribute(context, cxChild(parent, local), 'val');

/** Reports every child of `element` whose local name is not in `handled`. */
export function reportUnmodelled(
	context: ChartParseContext,
	element: XmlElement,
	handled: ReadonlySet<string>,
): void {
	for (const child of elements(element)) {
		if (handled.has(child.localName)) continue;
		context.issues.push({
			code: 'CHART_ELEMENT_NOT_MODELLED',
			message: `cx:${element.localName}/${child.prefix ? `${child.prefix}:` : ''}${child.localName} is not modelled; it is kept only in the source part.`,
		});
	}
}

/** A chartex child serialised as written (extension lists, caches), for round-trip. */
export function rawChild(parent: XmlElement | undefined, local: string): string | undefined {
	const element = cxChild(parent, local);
	return element ? buildXml(element) : undefined;
}

/** `cx:extLst` as written. */
export const extensionList = (parent: XmlElement | undefined) => rawChild(parent, 'extLst');

/** `cx:spPr` through the `drawingml` fill, line and effect readers; present for an empty element. */
export function shapeProperties(parent: XmlElement | undefined): ChartShapeProperties | undefined {
	const spPr = cxChild(parent, 'spPr');
	if (!spPr) return undefined;
	const out = readShapeFields(spPr);
	const source = innerXml(spPr);
	if (source) out.sourceXml = source;
	return out;
}

/** `cx:txPr` as a DrawingML text body. */
export const textProperties = (parent: XmlElement | undefined): ChartTextBody | undefined =>
	chartTextBody(cxChild(parent, 'txPr'));

/** Gridlines (`cx:majorGridlines`): present when written, with their shape. */
export function lines(parent: XmlElement | undefined, local: string): ChartLines | undefined {
	const element = cxChild(parent, local);
	if (!element) return undefined;
	const spPr = shapeProperties(element);
	return spPr ? { spPr } : {};
}

/** `cx:numFmt`. */
export function numberFormat(parent: XmlElement | undefined): ChartNumberFormat | undefined {
	const element = cxChild(parent, 'numFmt');
	const formatCode = attribute(element, 'formatCode');
	if (formatCode === undefined) return undefined;
	const linked = attribute(element, 'sourceLinked');
	return {
		formatCode,
		...(linked === '1' || linked === 'true' ? { sourceLinked: true } : {}),
		...(linked === '0' || linked === 'false' ? { sourceLinked: false } : {}),
	};
}

/** A formula element (`cx:f`, `cx:nf`). */
export function formula(parent: XmlElement | undefined, local: string): ChartExFormula | undefined {
	const element = cxChild(parent, local);
	if (!element) return undefined;
	const out: ChartExFormula = { formula: element.textContent ?? '' };
	const direction = attribute(element, 'dir');
	if (direction !== undefined) out.direction = direction;
	return out;
}
