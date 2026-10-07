import { NS, buildXml, elements, parseXml } from '../xml/index';
import { type ChartColorPalette } from './color-palettes';

export const CHART_COLOR_STYLE_NS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
export const CHART_COLOR_STYLE_REL =
	'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
export const CHART_COLOR_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartcolorstyle+xml';

/** Native chart color-style id, preserving unknown ids for callers to report. */
export function chartColorStyleId(xml: string): number | undefined {
	const root = parseXml(xml).documentElement;
	if (root.localName !== 'colorStyle' || root.namespaceURI !== CHART_COLOR_STYLE_NS)
		return undefined;
	const value = root.getAttribute('id');
	if (!value || !/^\d+$/.test(value)) return undefined;
	const id = Number(value);
	return Number.isInteger(id) && id >= 0 ? id : undefined;
}

/** Office palette metadata; retains unrelated attributes/extensions from an existing style. */
export function chartColorStyleXml(palette: ChartColorPalette, source?: string): string {
	const colors = palette.base.map((color) => `<a:schemeClr val="${color}"/>`).join('');
	const variations = palette.variations
		.map(
			(variation) =>
				`<cs:variation>${Object.entries(variation)
					.map(([name, value]) => `<a:${name} val="${value}"/>`)
					.join('')}</cs:variation>`,
		)
		.join('');
	const fresh = parseXml(
		`<cs:colorStyle xmlns:cs="${CHART_COLOR_STYLE_NS}" xmlns:a="${NS.a}" id="${palette.id}" meth="${palette.meth}">${colors}${variations}</cs:colorStyle>`,
	);
	if (!source) return buildXml(fresh);
	const doc = parseXml(source);
	const root = doc.documentElement;
	if (root.localName !== 'colorStyle' || root.namespaceURI !== CHART_COLOR_STYLE_NS)
		return buildXml(fresh);
	for (const child of elements(root)) {
		if (
			child.namespaceURI === NS.a ||
			(child.namespaceURI === CHART_COLOR_STYLE_NS && child.localName === 'variation')
		)
			root.removeChild(child);
	}
	for (const child of elements(fresh.documentElement))
		root.insertBefore(
			doc.importNode(child, true),
			elements(root).find(
				(node) => node.namespaceURI === CHART_COLOR_STYLE_NS && node.localName === 'extLst',
			) ?? null,
		);
	root.setAttribute('id', String(palette.id));
	root.setAttribute('meth', palette.meth);
	return buildXml(doc);
}
