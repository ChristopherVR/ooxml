// Shared helpers of the `c:chartSpace` DOM parser: chart-namespace lookups, typed `@val` readers
// that report malformed values instead of throwing, and the not-modelled child report.
import { NS, buildXml, elements, first, type XmlElement } from '../xml/index';
import { parseDrawingFill, parseDrawingLine } from '../drawingml/drawing-fill';
import { parseDrawingTextBody } from '../drawingml/drawing-text';
import type { DrawingTextBody } from '../drawingml/types';
import { readChartManualLayoutValues, type ChartManualLayout } from './manual-layout';
import type { ChartNumberFormat, ChartParseIssue, ChartShapeProperties } from './model-series';

/** Collects issues while a chart part is parsed. */
export interface ChartParseContext {
	issues: ChartParseIssue[];
}

/** The first chart-namespace child (lenient: un-namespaced test XML matches too). */
export const cChild = (parent: ParentNode | null | undefined, local: string) =>
	first(parent, local, NS.c);

/** An attribute that is written (`hasAttribute`), else `undefined`. */
export const attribute = (element: XmlElement | null | undefined, name: string) =>
	element?.hasAttribute(name) ? (element.getAttribute(name) ?? undefined) : undefined;

/** `@val` of the chart child `local`, as written. */
export const cVal = (parent: ParentNode | null | undefined, local: string) =>
	attribute(cChild(parent, local), 'val');

const path = (element: XmlElement) =>
	`c:${(element.parentNode as XmlElement | null)?.localName ?? ''}/c:${element.localName}`;

function malformed(context: ChartParseContext, element: XmlElement, value: string): void {
	context.issues.push({
		code: 'CHART_VALUE_INVALID',
		message: `${path(element)} has an unreadable value "${value}".`,
	});
}

/** A CT_Boolean child: absent -> undefined, no `@val` -> true (the schema default). */
export function cBool(
	context: ChartParseContext,
	parent: ParentNode | null | undefined,
	local: string,
): boolean | undefined {
	const element = cChild(parent, local);
	if (!element) return undefined;
	const raw = attribute(element, 'val');
	if (raw === undefined) return true;
	const value = raw.trim().toLowerCase();
	if (value === '1' || value === 'true') return true;
	if (value === '0' || value === 'false') return false;
	malformed(context, element, raw);
	return undefined;
}

/** A numeric `@val` child; malformed values are reported and dropped. */
export function cNumber(
	context: ChartParseContext,
	parent: ParentNode | null | undefined,
	local: string,
): number | undefined {
	const element = cChild(parent, local);
	const raw = attribute(element, 'val');
	if (!element || raw === undefined) return undefined;
	const value = raw.trim() ? Number(raw) : Number.NaN;
	if (Number.isFinite(value)) return value;
	malformed(context, element, raw);
	return undefined;
}

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
			message: `${path(child)} is not modelled; it is kept only in the source part.`,
		});
	}
}

/** `c:extLst` as written, for round-trip. */
export function extensionList(parent: XmlElement | undefined): string | undefined {
	const extLst = cChild(parent, 'extLst');
	return extLst ? buildXml(extLst) : undefined;
}

/** `c:spPr` through the `drawingml` fill, line and effect readers; present for an empty element. */
export function shapeProperties(parent: XmlElement | undefined): ChartShapeProperties | undefined {
	const spPr = cChild(parent, 'spPr');
	if (!spPr) return undefined;
	const out: ChartShapeProperties = {};
	const fill = parseDrawingFill(spPr);
	if (fill) out.fill = fill;
	const line = parseDrawingLine(first(spPr, 'ln', NS.a));
	if (line) out.line = line;
	const effects = first(spPr, 'effectLst', NS.a);
	if (effects) out.effectsXml = buildXml(effects);
	return out;
}

/** `c:txPr` as a DrawingML text body. */
export const textProperties = (parent: XmlElement | undefined): DrawingTextBody | undefined =>
	parseDrawingTextBody(cChild(parent, 'txPr'));

/** `c:numFmt`. */
export function numberFormat(parent: XmlElement | undefined): ChartNumberFormat | undefined {
	const element = cChild(parent, 'numFmt');
	const formatCode = attribute(element, 'formatCode');
	if (formatCode === undefined) return undefined;
	const linked = attribute(element, 'sourceLinked');
	return {
		formatCode,
		...(linked === '1' || linked === 'true' ? { sourceLinked: true } : {}),
		...(linked === '0' || linked === 'false' ? { sourceLinked: false } : {}),
	};
}

/** `c:layout/c:manualLayout`. */
export function manualLayout(parent: XmlElement | undefined): ChartManualLayout | undefined {
	const manual = cChild(cChild(parent, 'layout'), 'manualLayout');
	return manual ? readChartManualLayoutValues((name) => cVal(manual, name)) : undefined;
}

/** Assigns every defined entry of `values` to `target` (keeps optional properties absent). */
export function assignDefined<T extends object>(
	target: T,
	values: { [K in keyof T]?: T[K] | undefined },
): T {
	for (const [key, value] of Object.entries(values))
		if (value !== undefined) (target as Record<string, unknown>)[key] = value;
	return target;
}
