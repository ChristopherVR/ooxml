import { renderChartSvg } from 'ooxml-core/xlsx';
import type { VisioShape } from 'ooxml-core/visio';
import { parseSvgPreview } from '../ribbon/safe-svg';
import { svgElement } from './render-svg';

/** Pixels per inch the chart is laid out at before it is scaled into the shape's frame. */
const DPI = 96;

/**
 * An embedded chart object (Insert > Chart) drawn from its own workbook, so it stays sharp at any
 * zoom and follows resizing. The chart's SVG comes from trusted geometry and escaped text and is
 * parsed as a document with anything executable removed. Drawn like a picture: into the shape's
 * y-up frame, flipped back so it reads upright.
 */
export function renderEmbeddedChart(shape: VisioShape): SVGGElement | undefined {
	const object = shape.embeddedObject;
	if (!object) return undefined;
	const width = object.width ?? shape.width;
	const height = object.height ?? shape.height;
	if (!(width > 0) || !(height > 0)) return undefined;
	const markup = renderChartSvg(object.chart, Math.round(width * DPI), Math.round(height * DPI));
	const chart = parseSvgPreview(document, markup);
	if (!chart) return undefined;
	const node = svgElement('svg');
	for (const child of Array.from(chart.childNodes)) node.append(document.importNode(child, true));
	node.setAttribute(
		'viewBox',
		chart.getAttribute('viewBox') ?? `0 0 ${width * DPI} ${height * DPI}`,
	);
	node.setAttribute('width', String(width));
	node.setAttribute('height', String(height));
	node.setAttribute('preserveAspectRatio', 'none');
	node.setAttribute(
		'transform',
		`translate(${object.x ?? 0} ${(object.y ?? 0) + height}) scale(1 -1)`,
	);
	node.setAttribute('opacity', String(object.opacity ?? 1));
	const group = svgElement('g');
	group.dataset.embeddedChart = '';
	group.append(node);
	return group;
}
