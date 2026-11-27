/**
 * Drawing helpers for the portfolio's pictures, on top of the shared `Canvas`: ellipses, soft
 * shadows, gradients inside a box, and the small marks (wheat, leaf, drop, sun, arrow) the
 * project boards are composed from.
 */
import { Canvas, mix, type Rgb } from '../../lib/art';

type Point = readonly [number, number];

const clamp = (value: number) => Math.min(1, Math.max(0, value));

/** A filled ellipse, optionally rotated (radians), antialiased. */
export function ellipse(
	canvas: Canvas,
	cx: number,
	cy: number,
	rx: number,
	ry: number,
	color: Rgb,
	alpha = 1,
	rotation = 0,
): void {
	const reach = Math.max(rx, ry) + 1;
	const cos = Math.cos(rotation);
	const sin = Math.sin(rotation);
	const edge = Math.min(rx, ry);
	canvas.paint(
		{ x0: cx - reach, y0: cy - reach, x1: cx + reach, y1: cy + reach },
		color,
		(x, y) => {
			const dx = (x - cx) * cos + (y - cy) * sin;
			const dy = -(x - cx) * sin + (y - cy) * cos;
			return alpha * clamp((1 - Math.hypot(dx / rx, dy / ry)) * edge + 0.5);
		},
	);
}

/** A blurred elliptical shadow, darkest at its centre. */
export function shadow(
	canvas: Canvas,
	cx: number,
	cy: number,
	rx: number,
	ry: number,
	alpha: number,
	color: Rgb = [30, 24, 18],
): void {
	canvas.paint(
		{ x0: cx - rx * 2, y0: cy - ry * 2, x1: cx + rx * 2, y1: cy + ry * 2 },
		color,
		(x, y) => {
			const d = Math.hypot((x - cx) / rx, (y - cy) / ry);
			return alpha * Math.exp(-d * d * 1.6);
		},
	);
}

/** A rectangle filled with a vertical gradient. */
export function gradientRect(
	canvas: Canvas,
	x: number,
	y: number,
	w: number,
	h: number,
	top: Rgb,
	bottom: Rgb,
): void {
	canvas.paint(
		{ x0: x, y0: y, x1: x + w, y1: y + h },
		(_, py) => mix(top, bottom, (py - y) / h),
		(px, py) => (px >= x && px <= x + w && py >= y && py <= y + h ? 1 : 0),
	);
}

/** A lens (leaf) shape between two points, `bulge` being its half width. */
export function lens(
	canvas: Canvas,
	from: Point,
	to: Point,
	bulge: number,
	color: Rgb,
	alpha = 1,
): void {
	const dx = to[0] - from[0];
	const dy = to[1] - from[1];
	const length = Math.hypot(dx, dy) || 1;
	const nx = -dy / length;
	const ny = dx / length;
	const points: Point[] = [];
	const steps = 18;
	for (const side of [1, -1])
		for (let i = 0; i <= steps; i++) {
			const t = side === 1 ? i / steps : 1 - i / steps;
			const swell = Math.sin(Math.PI * t) * bulge * side;
			points.push([from[0] + dx * t + nx * swell, from[1] + dy * t + ny * swell]);
		}
	canvas.polygon(points, color, alpha);
}

/** A smooth open curve through points, drawn as short round-capped segments. */
export function curve(
	canvas: Canvas,
	points: readonly Point[],
	stroke: number,
	color: Rgb,
	alpha = 1,
): void {
	for (let i = 1; i < points.length; i++) {
		const a = points[i - 1]!;
		const b = points[i]!;
		canvas.line(a[0], a[1], b[0], b[1], stroke, color, alpha);
	}
}

/** An ear of wheat growing upwards from its base: a stem with paired grains and awns. */
export function wheat(
	canvas: Canvas,
	x: number,
	base: number,
	height: number,
	color: Rgb,
	scale = 1,
): void {
	canvas.line(x, base, x, base - height, 5 * scale, color);
	const grains = 7;
	const head = height * 0.56;
	for (let i = 0; i < grains; i++) {
		const y = base - height + head * (i / (grains - 1)) + 20 * scale;
		for (const side of [-1, 1]) {
			ellipse(canvas, x + side * 19 * scale, y, 26 * scale, 11 * scale, color, 1, side * -0.95);
			canvas.line(
				x + side * 30 * scale,
				y - 20 * scale,
				x + side * 52 * scale,
				y - 62 * scale,
				2.4 * scale,
				color,
			);
		}
	}
	ellipse(canvas, x, base - height - 6 * scale, 11 * scale, 27 * scale, color);
}

/** A water drop: a disc with a pointed top. */
export function drop(canvas: Canvas, cx: number, cy: number, radius: number, color: Rgb): void {
	canvas.circle(cx, cy + radius * 0.25, radius, color);
	canvas.polygon(
		[
			[cx, cy - radius * 1.45],
			[cx + radius * 0.86, cy - radius * 0.26],
			[cx - radius * 0.86, cy - radius * 0.26],
		],
		color,
	);
}

/** A sun: a disc with short rays. */
export function sun(canvas: Canvas, cx: number, cy: number, radius: number, color: Rgb): void {
	canvas.circle(cx, cy, radius * 0.52, color);
	for (let i = 0; i < 8; i++) {
		const angle = (Math.PI * 2 * i) / 8;
		canvas.line(
			cx + Math.cos(angle) * radius * 0.78,
			cy + Math.sin(angle) * radius * 0.78,
			cx + Math.cos(angle) * radius * 1.08,
			cy + Math.sin(angle) * radius * 1.08,
			radius * 0.17,
			color,
		);
	}
}

/** A straight arrow with a triangular head, pointing along `angle` (radians). */
export function arrow(
	canvas: Canvas,
	cx: number,
	cy: number,
	length: number,
	angle: number,
	stroke: number,
	color: Rgb,
): void {
	const ux = Math.cos(angle);
	const uy = Math.sin(angle);
	const tip: Point = [cx + (ux * length) / 2, cy + (uy * length) / 2];
	const tail: Point = [cx - (ux * length) / 2, cy - (uy * length) / 2];
	const neck: Point = [tip[0] - ux * stroke * 2.2, tip[1] - uy * stroke * 2.2];
	canvas.line(tail[0], tail[1], neck[0], neck[1], stroke, color);
	canvas.polygon(
		[
			tip,
			[neck[0] - uy * stroke * 1.9, neck[1] + ux * stroke * 1.9],
			[neck[0] + uy * stroke * 1.9, neck[1] - ux * stroke * 1.9],
		],
		color,
	);
}

/** A leaf on a short stalk, pointing along `angle` (radians). */
export function leaf(
	canvas: Canvas,
	cx: number,
	cy: number,
	size: number,
	angle: number,
	color: Rgb,
): void {
	const ux = Math.cos(angle);
	const uy = Math.sin(angle);
	lens(
		canvas,
		[cx - ux * size * 0.5, cy - uy * size * 0.5],
		[cx + ux * size * 0.62, cy + uy * size * 0.62],
		size * 0.34,
		color,
	);
	canvas.line(
		cx - ux * size * 0.5,
		cy - uy * size * 0.5,
		cx - ux * size * 0.82,
		cy - uy * size * 0.82,
		size * 0.09,
		color,
	);
}

/** Rows of bars standing in for lines of text. */
export function textLines(
	canvas: Canvas,
	x: number,
	y: number,
	widths: readonly number[],
	stroke: number,
	gap: number,
	color: Rgb,
	alpha = 1,
): void {
	widths.forEach((width, index) =>
		canvas.rect(x, y + index * gap, width, stroke, color, alpha, stroke / 2),
	);
}
