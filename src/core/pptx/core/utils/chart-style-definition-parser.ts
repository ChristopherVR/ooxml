/** Compatibility adapter: shared XML chart-style reader and formatting precedence. */
import { XMLBuilder } from 'fast-xml-parser';
import { CHART_COLOR_STYLE_NS } from '../../../chart/color-style';
import { readChartStyle } from '../../../chart/read-style';
import { CHART_STYLE_PARTS, resolveChartStyleDefinition } from '../../../chart/style-definition';
import { drawingColorXml } from '../../../drawingml/write-color';
import { NS, elements, parseXml } from '../../../xml/index';
import type { PptxChartStyleDefinition, XmlObject } from '../types';

interface XmlLookupLike {
	getChildByLocalName: (parent: XmlObject | undefined, name: string) => XmlObject | undefined;
}
type ResolveSchemeColor = (schemeClrNode: unknown) => string | undefined;
type ParseColor = (fillNode: XmlObject | undefined) => string | undefined;

const CS_CHILDREN = new Set(['defRPr', 'lnRef', 'fillRef', 'effectRef', 'fontRef', 'spPr']);
function normalize(value: unknown): unknown {
	if (Array.isArray(value)) return value.map(normalize);
	if (!value || typeof value !== 'object') return value;
	return Object.fromEntries(
		Object.entries(value).map(([key, item]) => {
			if (key.startsWith('@_') || key.startsWith('#')) return [key, item];
			const local = key.replace(/^.*:/u, '');
			return [`${CS_CHILDREN.has(local) ? 'cs' : 'a'}:${local}`, normalize(item)];
		}),
	);
}

export function parseChartStyleDefinition(
	styleRoot: XmlObject,
	xmlLookup: XmlLookupLike,
	resolveSchemeColor: ResolveSchemeColor,
	parseColor: ParseColor,
): PptxChartStyleDefinition | undefined {
	const parts: Record<string, unknown> = { '@_xmlns:cs': CHART_COLOR_STYLE_NS, '@_xmlns:a': NS.a };
	for (const name of CHART_STYLE_PARTS) {
		const node = xmlLookup.getChildByLocalName(styleRoot, name);
		if (node) parts[`cs:${name}`] = normalize(node);
	}
	const xml = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_' }).build({
		'cs:chartStyle': parts,
	});
	const style = readChartStyle(xml);
	if (!style) return undefined;
	return resolveChartStyleDefinition(style, (color) => {
		const element = parseXml(drawingColorXml(color)).documentElement;
		const raw: XmlObject = {};
		for (const attribute of Array.from(element.attributes))
			if (attribute.localName !== 'xmlns' && attribute.prefix !== 'xmlns')
				raw[`@_${attribute.localName}`] = attribute.value;
		for (const transform of elements(element))
			raw[`a:${transform.localName}`] = transform.hasAttribute('val')
				? { '@_val': transform.getAttribute('val') ?? '' }
				: {};
		const fill: XmlObject = {};
		const key = `a:${element.localName}` as `a:${string}`;
		fill[key] = raw;
		return color.kind === 'scheme' ? resolveSchemeColor(raw) : parseColor(fill);
	});
}
