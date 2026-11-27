/**
 * Procedural raster art for the samples: every picture in a sample is drawn here, pixel by pixel,
 * and encoded with the core's own PNG encoder. Nothing is downloaded or traced.
 */
import { encodePng } from '../../../src/core/drawingml/raster-encode';

export type Rgb = readonly [number, number, number];

export function hex(color: string): Rgb {
	const value = color.replace('#', '');
	return [
		Number.parseInt(value.slice(0, 2), 16),
		Number.parseInt(value.slice(2, 4), 16),
		Number.parseInt(value.slice(4, 6), 16),
	];
}

export function mix(a: Rgb, b: Rgb, t: number): Rgb {
	const k = Math.min(1, Math.max(0, t));
	return [a[0] + (b[0] - a[0]) * k, a[1] + (b[1] - a[1]) * k, a[2] + (b[2] - a[2]) * k];
}

interface Box {
	x0: number;
	y0: number;
	x1: number;
	y1: number;
}

/** An RGB canvas with the few primitives the sample art needs. Coordinates are pixels. */
export class Canvas {
	readonly pixels: Uint8Array;
	constructor(
		readonly width: number,
		readonly height: number,
		background: Rgb = [255, 255, 255],
	) {
		this.pixels = new Uint8Array(width * height * 3);
		this.fill(() => background);
	}

	/** Paints every pixel from `shade(u, v)`, with u and v in [0, 1]. */
	fill(shade: (u: number, v: number) => Rgb): this {
		for (let y = 0; y < this.height; y++)
			for (let x = 0; x < this.width; x++)
				this.set(x, y, shade(x / (this.width - 1 || 1), y / (this.height - 1 || 1)), 1);
		return this;
	}

	set(x: number, y: number, color: Rgb, alpha: number): void {
		if (x < 0 || y < 0 || x >= this.width || y >= this.height || alpha <= 0) return;
		const at = (Math.floor(y) * this.width + Math.floor(x)) * 3;
		const a = Math.min(1, alpha);
		for (let c = 0; c < 3; c++)
			this.pixels[at + c] = Math.round(this.pixels[at + c]! * (1 - a) + color[c]! * a);
	}

	/** Blends `color` wherever `coverage(x, y)` (0 to 1) is positive, inside the given box. */
	paint(
		box: Box,
		color: Rgb | ((x: number, y: number) => Rgb),
		coverage: (x: number, y: number) => number,
	): this {
		const x0 = Math.max(0, Math.floor(box.x0));
		const y0 = Math.max(0, Math.floor(box.y0));
		const x1 = Math.min(this.width - 1, Math.ceil(box.x1));
		const y1 = Math.min(this.height - 1, Math.ceil(box.y1));
		for (let y = y0; y <= y1; y++)
			for (let x = x0; x <= x1; x++) {
				const a = coverage(x + 0.5, y + 0.5);
				if (a > 0) this.set(x, y, typeof color === 'function' ? color(x, y) : color, a);
			}
		return this;
	}

	/** An antialiased disc. */
	circle(cx: number, cy: number, radius: number, color: Rgb, alpha = 1): this {
		return this.paint(
			{ x0: cx - radius - 1, y0: cy - radius - 1, x1: cx + radius + 1, y1: cy + radius + 1 },
			color,
			(x, y) => alpha * Math.min(1, Math.max(0, radius - Math.hypot(x - cx, y - cy) + 0.5)),
		);
	}

	/** An antialiased ring of the given stroke width. */
	ring(cx: number, cy: number, radius: number, stroke: number, color: Rgb, alpha = 1): this {
		const outer = radius + stroke / 2;
		return this.paint(
			{ x0: cx - outer - 1, y0: cy - outer - 1, x1: cx + outer + 1, y1: cy + outer + 1 },
			color,
			(x, y) =>
				alpha *
				Math.min(1, Math.max(0, stroke / 2 - Math.abs(Math.hypot(x - cx, y - cy) - radius) + 0.5)),
		);
	}

	/** A rectangle with optional rounded corners. */
	rect(x: number, y: number, w: number, h: number, color: Rgb, alpha = 1, radius = 0): this {
		return this.paint({ x0: x, y0: y, x1: x + w, y1: y + h }, color, (px, py) => {
			if (px < x || px > x + w || py < y || py > y + h) return 0;
			if (radius <= 0) return alpha;
			const dx = Math.max(x + radius - px, 0, px - (x + w - radius));
			const dy = Math.max(y + radius - py, 0, py - (y + h - radius));
			return alpha * Math.min(1, Math.max(0, radius - Math.hypot(dx, dy) + 0.5));
		});
	}

	/** A line segment with round ends. */
	line(
		x0: number,
		y0: number,
		x1: number,
		y1: number,
		stroke: number,
		color: Rgb,
		alpha = 1,
	): this {
		const dx = x1 - x0;
		const dy = y1 - y0;
		const length = dx * dx + dy * dy || 1;
		return this.paint(
			{
				x0: Math.min(x0, x1) - stroke,
				y0: Math.min(y0, y1) - stroke,
				x1: Math.max(x0, x1) + stroke,
				y1: Math.max(y0, y1) + stroke,
			},
			color,
			(x, y) => {
				const t = Math.min(1, Math.max(0, ((x - x0) * dx + (y - y0) * dy) / length));
				const distance = Math.hypot(x - (x0 + t * dx), y - (y0 + t * dy));
				return alpha * Math.min(1, Math.max(0, stroke / 2 - distance + 0.5));
			},
		);
	}

	/** A filled polygon (even-odd rule), antialiased with 3x3 supersampling. */
	polygon(points: readonly (readonly [number, number])[], color: Rgb, alpha = 1): this {
		const xs = points.map((p) => p[0]);
		const ys = points.map((p) => p[1]);
		const inside = (x: number, y: number) => {
			let hit = false;
			for (let i = 0, j = points.length - 1; i < points.length; j = i++) {
				const [xi, yi] = points[i]!;
				const [xj, yj] = points[j]!;
				if (yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) hit = !hit;
			}
			return hit;
		};
		return this.paint(
			{ x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) },
			color,
			(x, y) => {
				let hits = 0;
				for (const ox of [-1 / 3, 0, 1 / 3])
					for (const oy of [-1 / 3, 0, 1 / 3]) if (inside(x + ox, y + oy)) hits++;
				return (alpha * hits) / 9;
			},
		);
	}

	toPng(): Uint8Array {
		return encodePng(this.width, this.height, this.pixels);
	}
}

/** A smooth two-dimensional value noise in [0, 1], for organic gradients. */
export function valueNoise(seed: number): (x: number, y: number) => number {
	const hash = (x: number, y: number) => {
		let h = (Math.imul(x, 374761393) + Math.imul(y, 668265263) + Math.imul(seed, 2246822519)) | 0;
		h = Math.imul(h ^ (h >>> 13), 1274126177);
		return ((h ^ (h >>> 16)) >>> 0) / 4294967296;
	};
	const smooth = (t: number) => t * t * (3 - 2 * t);
	return (x, y) => {
		const xi = Math.floor(x);
		const yi = Math.floor(y);
		const tx = smooth(x - xi);
		const ty = smooth(y - yi);
		const top = hash(xi, yi) * (1 - tx) + hash(xi + 1, yi) * tx;
		const bottom = hash(xi, yi + 1) * (1 - tx) + hash(xi + 1, yi + 1) * tx;
		return top * (1 - ty) + bottom * ty;
	};
}
