import { buildChartGradientDef, type ChartGradientFill } from 'ooxml-core/chart';

/** Preview a gesture in the rendered SVG; the workbook is edited once on release. */
export function createChartGradientPreview(
	root: ShadowRoot,
	drawing: number,
	series: number,
	gradient: ChartGradientFill,
) {
	const nodes = Array.from(
		root.querySelectorAll<SVGElement>(
			`.xg-obj[data-index="${drawing}"] :is(linearGradient,radialGradient,pattern)[id$="-s${series}"]`,
		),
	);
	const originals = nodes.map((node) => ({
		node,
		children: Array.from(node.childNodes).map((child) => child.cloneNode(true)),
	}));
	const paint = (fill: ChartGradientFill) => {
		const def = buildChartGradientDef('preview', fill);
		for (const node of nodes) {
			if (!node.isConnected) continue;
			if (def.kind === 'rectPath') {
				const image = node.ownerDocument.createElementNS('http://www.w3.org/2000/svg', 'image');
				image.setAttribute('href', def.href);
				image.setAttribute('width', '1');
				image.setAttribute('height', '1');
				image.setAttribute('preserveAspectRatio', 'none');
				node.replaceChildren(image);
				continue;
			}
			node.replaceChildren(
				...def.stops.map((stop) => {
					const element = node.ownerDocument.createElementNS('http://www.w3.org/2000/svg', 'stop');
					element.setAttribute('offset', String(stop.offset));
					element.setAttribute('stop-color', stop.color);
					element.setAttribute('stop-opacity', String(stop.opacity ?? 1));
					return element;
				}),
			);
		}
	};
	return {
		paint,
		position(index: number, position: number) {
			paint({
				...gradient,
				stops: gradient.stops.map((stop, i) => (i === index ? { ...stop, position } : stop)),
			});
		},
		restore() {
			for (const { node, children } of originals)
				if (node.isConnected) node.replaceChildren(...children);
		},
	};
}
