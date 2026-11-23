/** SVG path data to absolute move/line/cubic/close segments (PDF writer, 3-D projection). DOM-free. */

export type PathSegment =
	| { t: 'M'; x: number; y: number }
	| { t: 'L'; x: number; y: number }
	| { t: 'C'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { t: 'Z' };

export const MAX_PATH_CHARS = 2_000_000;
export const MAX_PATH_SEGMENTS = 200_000;

const COUNTS: Record<string, number> = {
	m: 2,
	l: 2,
	h: 1,
	v: 1,
	c: 6,
	s: 4,
	q: 4,
	t: 2,
	a: 7,
	z: 0,
};

/** Parse path data. Returns `undefined` if it is malformed or exceeds the bounds. */
export function parsePath(d: string): PathSegment[] | undefined {
	if (d.length > MAX_PATH_CHARS) return undefined;
	const tokens = d.match(/[a-zA-Z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g);
	if (!tokens) return [];
	const out: PathSegment[] = [];
	let x = 0,
		y = 0,
		startX = 0,
		startY = 0,
		lastC: [number, number] | undefined,
		lastQ: [number, number] | undefined;
	let command = '';
	let index = 0;
	const read = () => Number(tokens[index++]);
	while (index < tokens.length) {
		if (out.length > MAX_PATH_SEGMENTS) return undefined;
		const token = tokens[index]!;
		if (/[a-zA-Z]/.test(token)) {
			command = token;
			index++;
		} else if (!command) return undefined;
		else if (command === 'M') command = 'L';
		else if (command === 'm') command = 'l';
		const lower = command.toLowerCase();
		const need = COUNTS[lower];
		if (need === undefined) return undefined;
		if (index + need > tokens.length) return undefined;
		const relative = command === lower;
		const ox = relative ? x : 0,
			oy = relative ? y : 0;
		let nextC: [number, number] | undefined, nextQ: [number, number] | undefined;
		switch (lower) {
			case 'm':
				x = ox + read();
				y = oy + read();
				startX = x;
				startY = y;
				out.push({ t: 'M', x, y });
				break;
			case 'l':
				x = ox + read();
				y = oy + read();
				out.push({ t: 'L', x, y });
				break;
			case 'h':
				x = (relative ? x : 0) + read();
				out.push({ t: 'L', x, y });
				break;
			case 'v':
				y = (relative ? y : 0) + read();
				out.push({ t: 'L', x, y });
				break;
			case 'c': {
				const x1 = ox + read(),
					y1 = oy + read(),
					x2 = ox + read(),
					y2 = oy + read();
				out.push({ t: 'C', x1, y1, x2, y2, x: (x = ox + read()), y: (y = oy + read()) });
				nextC = [x2, y2];
				break;
			}
			case 's': {
				const x1 = lastC ? 2 * x - lastC[0] : x,
					y1 = lastC ? 2 * y - lastC[1] : y;
				const x2 = ox + read(),
					y2 = oy + read();
				out.push({ t: 'C', x1, y1, x2, y2, x: (x = ox + read()), y: (y = oy + read()) });
				nextC = [x2, y2];
				break;
			}
			case 'q':
			case 't': {
				let qx: number, qy: number;
				if (lower === 'q') {
					qx = ox + read();
					qy = oy + read();
				} else {
					qx = lastQ ? 2 * x - lastQ[0] : x;
					qy = lastQ ? 2 * y - lastQ[1] : y;
				}
				const ex = ox + read(),
					ey = oy + read();
				out.push({
					t: 'C',
					x1: x + (2 / 3) * (qx - x),
					y1: y + (2 / 3) * (qy - y),
					x2: ex + (2 / 3) * (qx - ex),
					y2: ey + (2 / 3) * (qy - ey),
					x: ex,
					y: ey,
				});
				x = ex;
				y = ey;
				nextQ = [qx, qy];
				break;
			}
			case 'a': {
				const rx = read(),
					ry = read(),
					rotation = read(),
					large = read(),
					sweep = read();
				const ex = ox + read(),
					ey = oy + read();
				arc(out, x, y, rx, ry, rotation, large !== 0, sweep !== 0, ex, ey);
				x = ex;
				y = ey;
				break;
			}
			case 'z':
				out.push({ t: 'Z' });
				x = startX;
				y = startY;
				break;
		}
		lastC = nextC;
		lastQ = nextQ;
		if (lower === 'z') command = '';
	}
	return out;
}

/** Elliptical arc as cubic segments (SVG implementation notes, F.6). */
function arc(
	out: PathSegment[],
	x0: number,
	y0: number,
	rxIn: number,
	ryIn: number,
	degrees: number,
	large: boolean,
	sweep: boolean,
	x: number,
	y: number,
): void {
	let rx = Math.abs(rxIn),
		ry = Math.abs(ryIn);
	if ((x0 === x && y0 === y) || rx === 0 || ry === 0) {
		out.push({ t: 'L', x, y });
		return;
	}
	const phi = (degrees * Math.PI) / 180,
		cos = Math.cos(phi),
		sin = Math.sin(phi);
	const dx = (x0 - x) / 2,
		dy = (y0 - y) / 2;
	const x1 = cos * dx + sin * dy,
		y1 = -sin * dx + cos * dy;
	const lambda = (x1 * x1) / (rx * rx) + (y1 * y1) / (ry * ry);
	if (lambda > 1) {
		rx *= Math.sqrt(lambda);
		ry *= Math.sqrt(lambda);
	}
	const numerator = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1;
	const denominator = rx * rx * y1 * y1 + ry * ry * x1 * x1;
	const factor =
		(large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, numerator / (denominator || 1)));
	const cx1 = (factor * rx * y1) / ry,
		cy1 = (-factor * ry * x1) / rx;
	const cx = cos * cx1 - sin * cy1 + (x0 + x) / 2,
		cy = sin * cx1 + cos * cy1 + (y0 + y) / 2;
	const angle = (ux: number, uy: number, vx: number, vy: number) =>
		Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
	const start = angle(1, 0, (x1 - cx1) / rx, (y1 - cy1) / ry);
	let delta = angle((x1 - cx1) / rx, (y1 - cy1) / ry, (-x1 - cx1) / rx, (-y1 - cy1) / ry);
	if (!sweep && delta > 0) delta -= 2 * Math.PI;
	else if (sweep && delta < 0) delta += 2 * Math.PI;
	const pieces = Math.max(1, Math.ceil(Math.abs(delta) / (Math.PI / 2) - 1e-9));
	const step = delta / pieces;
	const k = (4 / 3) * Math.tan(step / 4);
	const point = (t: number): [number, number, number, number] => {
		const c = Math.cos(t),
			s = Math.sin(t);
		return [
			cx + rx * c * cos - ry * s * sin,
			cy + rx * c * sin + ry * s * cos,
			-rx * s * cos - ry * c * sin,
			-rx * s * sin + ry * c * cos,
		];
	};
	for (let piece = 0; piece < pieces; piece++) {
		const a = start + piece * step,
			b = a + step;
		const [px, py, dpx, dpy] = point(a);
		const [qx, qy, dqx, dqy] = point(b);
		out.push({
			t: 'C',
			x1: px + k * dpx,
			y1: py + k * dpy,
			x2: qx - k * dqx,
			y2: qy - k * dqy,
			x: piece === pieces - 1 ? x : qx,
			y: piece === pieces - 1 ? y : qy,
		});
	}
}
