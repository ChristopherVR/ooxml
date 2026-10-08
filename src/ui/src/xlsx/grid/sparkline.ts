// Paints the core's sparkline descriptor (a unit box) as an SVG filling the host cell, behind the
// cell's text and borders. All geometry and colours come from `sparklineView`; this only scales
// the box to the cell's pixel size at the current zoom.
import type { SparklineView } from 'ooxml-core/xlsx';

const SVG = 'http://www.w3.org/2000/svg';

const fixed = (n: number): string => String(Math.round(n * 100) / 100);

function svgElement(doc: Document, tag: string, attrs: Record<string, string>): Element {
	const node = doc.createElementNS(SVG, tag);
	for (const [name, value] of Object.entries(attrs)) node.setAttribute(name, value);
	return node;
}

/** Builds the `<svg class="xg-spark">` for a `w` x `h` px cell at `zoom` percent. */
export function sparklineSvg(
	doc: Document,
	view: SparklineView,
	w: number,
	h: number,
	zoom: number,
): Element {
	const scale = zoom / 100;
	const stroke = Math.max(0.5, view.lineWeightPt * (96 / 72) * scale);
	const radius = view.markers.length ? Math.max(1.5 * scale, stroke * 1.25) : 0;
	const padX = 2 * scale + (view.type === 'line' ? Math.max(radius, stroke / 2) : 0);
	const padY = 2 * scale + (view.type === 'line' ? Math.max(radius, stroke / 2) : 0);
	const innerW = Math.max(0, w - 2 * padX);
	const innerH = Math.max(0, h - 2 * padY);
	const px = (x: number) => padX + x * innerW;
	const py = (y: number) => padY + y * innerH;
	const svg = svgElement(doc, 'svg', {
		class: 'xg-spark',
		width: fixed(w),
		height: fixed(h),
		'aria-hidden': 'true',
	});
	if (view.axis)
		svg.append(
			svgElement(doc, 'line', {
				x1: fixed(padX),
				x2: fixed(padX + innerW),
				y1: fixed(py(view.axis.y)),
				y2: fixed(py(view.axis.y)),
				stroke: view.axis.color,
				'stroke-width': fixed(Math.max(1, scale)),
				'shape-rendering': 'crispEdges',
			}),
		);
	for (const column of view.columns) {
		// A column at the edge of its scale keeps a one-pixel sliver, as Excel draws it.
		const height = Math.max(1, column.h * innerH);
		const top = column.h * innerH < 1 && column.y >= 0.5 ? py(column.y) - 1 : py(column.y);
		svg.append(
			svgElement(doc, 'rect', {
				x: fixed(px(column.x)),
				y: fixed(top),
				width: fixed(Math.max(1, column.w * innerW)),
				height: fixed(height),
				fill: column.color,
			}),
		);
	}
	for (const line of view.lines) {
		if (line.length === 1) continue;
		svg.append(
			svgElement(doc, 'polyline', {
				points: line.map((p) => `${fixed(px(p.x))},${fixed(py(p.y))}`).join(' '),
				fill: 'none',
				stroke: view.color,
				'stroke-width': fixed(stroke),
				'stroke-linejoin': 'round',
				'stroke-linecap': 'round',
			}),
		);
	}
	for (const marker of view.markers)
		svg.append(
			svgElement(doc, 'circle', {
				cx: fixed(px(marker.x)),
				cy: fixed(py(marker.y)),
				r: fixed(radius),
				fill: marker.color,
			}),
		);
	return svg;
}
