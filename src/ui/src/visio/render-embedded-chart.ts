import { renderChartSvg } from 'ooxml-core/chart';
import type { VisioShape } from 'ooxml-core/visio';
import { parseSvgPreview } from '../ribbon/safe-svg';
import { svgElement } from './render-svg';

/** Pixels per inch the chart is laid out at before it is scaled into the shape's frame. */
const DPI = 96;

const MAP_APPROXIMATION_LABEL =
	'Map outlines here are approximate. Office draws a new map only after it adds map data through its online service.';

/**
 * An embedded chart object (Insert > Chart) drawn from its own workbook, so it stays sharp at any
 * zoom and follows resizing. Supports both classic `c:chartSpace` charts and Office 2016+ ChartEx
 * (`cx:chartSpace`) charts via the shared SVG painter. The chart's SVG comes from trusted geometry
 * and escaped text and is parsed as a document with anything executable removed. Drawn like a
 * picture: into the shape's y-up frame, flipped back so it reads upright.
 */
export function renderEmbeddedChart(shape: VisioShape): SVGGElement | undefined {
	const object = shape.embeddedObject;
	if (!object) return undefined;
	const width = object.width ?? shape.width;
	const height = object.height ?? shape.height;
	if (!(width > 0) || !(height > 0)) return undefined;
	// The shape's own alt text names the chart, and the painter's `<desc>` reads its data out; the
	// copy below keeps them, so it must also keep `role="img"` for a reader to use either.
	const altText = shape.alternativeText || shape.accessibilityTitle || '';
	const markup = renderChartSvg(object.chart, Math.round(width * DPI), Math.round(height * DPI), {
		...(altText ? { altText } : {}),
		...(object.chart.type === 'regionMap'
			? { mapApproximationLabel: MAP_APPROXIMATION_LABEL }
			: {}),
	});
	const chart = parseSvgPreview(document, markup);
	if (!chart) return undefined;
	const node = svgElement('svg');
	node.setAttribute('role', 'img');
	const fontFamily = chart.getAttribute('font-family');
	if (fontFamily) node.setAttribute('font-family', fontFamily);
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
	group.dataset.chartType = object.chart.type;
	if (object.chartEx) group.dataset.embeddedChartEx = '';
	group.append(node);
	return group;
}
