/** SVG basic shapes (`circle`, `ellipse`, `line`, `polyline`, `polygon`) as path data. */
import { number } from './pdf-svg-style';

/** Bezier constant for a quarter ellipse. */
const K = 0.5522847498;

function ellipse(cx: number, cy: number, rx: number, ry: number): string | undefined {
	if (!(rx > 0 && ry > 0)) return undefined;
	const ox = rx * K,
		oy = ry * K;
	return [
		`M ${cx + rx} ${cy}`,
		`C ${cx + rx} ${cy + oy} ${cx + ox} ${cy + ry} ${cx} ${cy + ry}`,
		`C ${cx - ox} ${cy + ry} ${cx - rx} ${cy + oy} ${cx - rx} ${cy}`,
		`C ${cx - rx} ${cy - oy} ${cx - ox} ${cy - ry} ${cx} ${cy - ry}`,
		`C ${cx + ox} ${cy - ry} ${cx + rx} ${cy - oy} ${cx + rx} ${cy}`,
		'Z',
	].join(' ');
}

/** The path data an SVG basic shape draws, or `undefined` when it draws nothing. */
export function basicShapePath(element: Element): string | undefined {
	const attr = (name: string) => number(element.getAttribute(name));
	switch (element.localName) {
		case 'circle':
			return ellipse(attr('cx'), attr('cy'), attr('r'), attr('r'));
		case 'ellipse':
			return ellipse(attr('cx'), attr('cy'), attr('rx'), attr('ry'));
		case 'line':
			return `M ${attr('x1')} ${attr('y1')} L ${attr('x2')} ${attr('y2')}`;
		case 'polyline':
		case 'polygon': {
			const values = (element.getAttribute('points') ?? '')
				.trim()
				.split(/[\s,]+/)
				.filter(Boolean)
				.map(Number);
			if (values.length < 4 || values.some((v) => !Number.isFinite(v))) return undefined;
			const points: string[] = [];
			for (let i = 0; i + 1 < values.length; i += 2)
				points.push(`${i ? 'L' : 'M'} ${values[i]} ${values[i + 1]}`);
			return points.join(' ') + (element.localName === 'polygon' ? ' Z' : '');
		}
		default:
			return undefined;
	}
}
