// The editing surface's chart: the painted SVG (as a ProseMirror DOM spec, so no markup string is
// ever parsed into the live document) sized to the drawing extent, or the labelled frame for a
// family the painter does not draw. Read-only: the chart is an atom whose part is preserved.
import type { DOMOutputSpec, Node as ProseMirrorNode } from 'prosemirror-model';
import { isElement, parseXml, type XmlElement } from '../../xml/index';
import type { DocxChart } from '../chart';
import { docxChartPaint } from '../chart-paint';
import type { ThemeCatalog } from '../theme-model';
import { inlineRunPropertiesDomAttrs } from './inline-run-properties';

const SVG_NS = 'http://www.w3.org/2000/svg';

function elementSpec(element: XmlElement, root: boolean): DOMOutputSpec {
	const attrs: Record<string, string> = {};
	for (const attr of Array.from(element.attributes)) {
		// Namespace declarations come from the element namespace; event handlers are never copied.
		if (attr.name === 'xmlns' || attr.name.startsWith('xmlns:') || /^on/i.test(attr.name)) continue;
		attrs[attr.name] = attr.value;
	}
	// Text children are plain strings, which ProseMirror renders as text nodes.
	const children: unknown[] = [];
	for (const child of Array.from(element.childNodes)) {
		if (isElement(child)) children.push(elementSpec(child, false));
		else if (child.nodeType === 3 && child.nodeValue) children.push(child.nodeValue);
	}
	const tag = root ? `${SVG_NS} ${element.localName}` : element.localName;
	return [tag, attrs, ...children] as unknown as DOMOutputSpec;
}

/** An SVG document as a ProseMirror DOM spec in the SVG namespace. */
export function svgMarkupSpec(svg: string): DOMOutputSpec {
	return elementSpec(parseXml(svg, { label: 'Chart SVG' }).documentElement, true);
}

/** The chart model carried on an image node (`attrs.chart`, JSON), if any. */
export function nodeChart(node: ProseMirrorNode): DocxChart | undefined {
	const value: unknown = node.attrs.chart;
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as DocxChart;
	} catch {
		return undefined;
	}
}

/** The DOM spec of a chart image node, painted with `theme` (the Office theme when absent). */
export function chartNodeSpec(node: ProseMirrorNode, theme?: ThemeCatalog): DOMOutputSpec {
	const width = Number(node.attrs.widthPx) || 96;
	const height = Number(node.attrs.heightPx) || 96;
	const paint = docxChartPaint(nodeChart(node), width, height, theme);
	const style = `width:${width}px;height:${height}px`;
	if (!paint.svg)
		return [
			'span',
			{
				'data-docx-image-placeholder': '1',
				'data-docx-chart': 'frame',
				...inlineRunPropertiesDomAttrs(node),
				class: 'dve-image-placeholder',
				style,
			},
			paint.label,
		];
	return [
		'span',
		{
			'data-docx-chart': '1',
			...inlineRunPropertiesDomAttrs(node),
			class: 'dve-chart',
			style,
			title: paint.label,
		},
		svgMarkupSpec(paint.svg),
	];
}
