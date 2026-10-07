import { parseDrawingColorIn } from '../diagram/drawing-color';
import { parseDrawingFill, parseDrawingLine } from '../diagram/drawing-fill';
import { NS, buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';
import { CHART_COLOR_STYLE_NS } from './color-style';
import type {
	ChartStyleDefinition,
	ChartStyleEntry,
	ChartStyleReference,
} from './style-definition';

const child = (parent: XmlElement | undefined, name: string) =>
	first(parent, name, CHART_COLOR_STYLE_NS);

function reference(node: XmlElement | undefined): ChartStyleReference | undefined {
	if (!node) return undefined;
	const out: ChartStyleReference = { index: node.getAttribute('idx') ?? '' };
	const color = parseDrawingColorIn(node);
	if (color) out.color = color;
	return out;
}

function entry(node: XmlElement): ChartStyleEntry {
	const out: ChartStyleEntry = { sourceXml: buildXml(node) };
	const text = child(node, 'defRPr');
	const size = Number(text?.getAttribute('sz') ?? NaN);
	if (Number.isFinite(size)) out.fontSize = size / 100;
	for (const [attribute, key] of [
		['b', 'bold'],
		['i', 'italic'],
	] as const) {
		const value = text?.getAttribute(attribute);
		if (value !== null && value !== undefined) out[key] = value === '1' || value === 'true';
	}
	const typeface = first(text, 'latin', NS.a)?.getAttribute('typeface');
	if (typeface) out.typeface = typeface;
	const color = parseDrawingColorIn(first(text, 'solidFill', NS.a));
	if (color) out.textColor = color;
	for (const key of ['fontRef', 'lineRef', 'fillRef', 'effectRef'] as const) {
		const ref = reference(child(node, key === 'lineRef' ? 'lnRef' : key));
		if (ref) out[key] = ref;
	}
	const properties = child(node, 'spPr');
	const fill = parseDrawingFill(properties);
	const line = parseDrawingLine(first(properties, 'ln', NS.a));
	if (fill) out.fill = fill;
	if (line) out.line = line;
	return out;
}

/** Read once against the shared XML/DrawingML model, without flattening theme choices. */
export function readChartStyle(xml: string): ChartStyleDefinition | undefined {
	const root = parseXml(xml, { label: 'Office chart style' }).documentElement;
	if (root.namespaceURI !== CHART_COLOR_STYLE_NS || root.localName !== 'chartStyle')
		return undefined;
	const out: ChartStyleDefinition = { entries: {}, sourceXml: xml };
	const id = root.getAttribute('id');
	if (id && /^\d+$/.test(id) && Number.isSafeInteger(Number(id))) out.id = Number(id);
	out.entries = Object.fromEntries(
		elements(root)
			.filter((node) => node.namespaceURI === CHART_COLOR_STYLE_NS && node.localName !== 'extLst')
			.map((node) => [node.localName, entry(node)]),
	);
	return out;
}
