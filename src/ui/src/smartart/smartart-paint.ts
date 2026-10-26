import { buildRectPathGradientSvg, resolveDrawingColor } from 'ooxml-core/diagram';
import type { DrawingFill } from 'ooxml-core/drawingml';
import { colorToCss, type SchemeColors, type SmartArtRenderReport } from './smartart-svg';

const NS = 'http://www.w3.org/2000/svg';
let serial = 0;

/** Each SVG owns its paint servers; ids remain distinct when several diagrams share a document. */
export function diagramFill(
	doc: Document,
	defs: SVGDefsElement,
	fill: DrawingFill | undefined,
	scheme: SchemeColors,
	report: SmartArtRenderReport,
): string {
	if (!fill || fill.kind === 'none') return 'none';
	if (fill.kind === 'solid') return colorToCss(fill.color, scheme, report) ?? 'none';
	if (fill.kind !== 'gradient') {
		if (!report.approximatedFills.includes(fill.kind)) report.approximatedFills.push(fill.kind);
		return '#d1d5db';
	}
	const id = `office-smartart-paint-${++serial}`;
	if (fill.path === 'rect') {
		const stops = fill.stops.flatMap((stop) => {
			const color = resolveDrawingColor(stop.color, { scheme: (name) => scheme[name] });
			return color ? [{ position: stop.position, color: color.hex, opacity: color.alpha }] : [];
		});
		const pattern = doc.createElementNS(NS, 'pattern');
		pattern.id = id;
		pattern.setAttribute('width', '1');
		pattern.setAttribute('height', '1');
		pattern.setAttribute('patternContentUnits', 'objectBoundingBox');
		const image = doc.createElementNS(NS, 'image');
		image.setAttribute('width', '1');
		image.setAttribute('height', '1');
		image.setAttribute('preserveAspectRatio', 'none');
		image.setAttribute(
			'href',
			`data:image/svg+xml,${encodeURIComponent(buildRectPathGradientSvg(stops, undefined, fill.fillToRect, { bands: 40, independentOpacity: true }))}`,
		);
		pattern.append(image);
		defs.append(pattern);
		return `url(#${id})`;
	}
	const radial = Boolean(fill.path);
	const gradient = doc.createElementNS(NS, radial ? 'radialGradient' : 'linearGradient');
	gradient.id = id;
	if (radial) {
		const rect = fill.fillToRect;
		gradient.setAttribute('cx', String(rect ? (rect.l + 1 - rect.r) / 2 : 0.5));
		gradient.setAttribute('cy', String(rect ? (rect.t + 1 - rect.b) / 2 : 0.5));
		gradient.setAttribute('r', '0.5');
		if (fill.path !== 'circle' && !report.approximatedFills.includes(`gradient:${fill.path}`))
			report.approximatedFills.push(`gradient:${fill.path}`);
	} else {
		const radians = ((fill.angle ?? 0) * Math.PI) / 180;
		const dx = Math.cos(radians) / 2;
		const dy = Math.sin(radians) / 2;
		gradient.setAttribute('x1', String(0.5 - dx));
		gradient.setAttribute('y1', String(0.5 - dy));
		gradient.setAttribute('x2', String(0.5 + dx));
		gradient.setAttribute('y2', String(0.5 + dy));
	}
	for (const stop of fill.stops) {
		const element = doc.createElementNS(NS, 'stop');
		element.setAttribute('offset', `${stop.position}%`);
		element.setAttribute('stop-color', colorToCss(stop.color, scheme, report) ?? '#808080');
		gradient.append(element);
	}
	defs.append(gradient);
	return `url(#${id})`;
}
