/**
 * The inverse discrete wavelet transform (ITU-T T.800 Annex F): 2D_SR over the resolution levels,
 * each an interleave of the four sub-bands, then the one-dimensional reconstruction of every row
 * and every column with periodic symmetric extension, by the reversible 5-3 or the irreversible
 * 9-7 lifting steps. Sample positions keep their parity on the reference grid (F.3.7).
 */
import type { Budget } from './errors';
import type { TileComponent } from './structure';

const PAD = 4;
const ALPHA = -1.586134342059924;
const BETA = -0.052980118572961;
const GAMMA = 0.882911075530934;
const DELTA = 0.443506852043971;
const K = 1.230174104914001;

let line = new Float64Array(64);

/** Reconstructs `count` samples at `x[PAD..]`, the first at absolute position `i0`. */
function filter(x: Float64Array, i0: number, count: number, reversible: boolean): void {
	if (count === 1) {
		if (i0 & 1) x[PAD] = x[PAD]! / 2;
		return;
	}
	// Periodic symmetric extension (F.3.7).
	for (let k = 1; k <= PAD; k++) {
		const period = 2 * (count - 1);
		const mirror = (offset: number) => {
			let m = ((offset % period) + period) % period;
			if (m >= count) m = period - m;
			return m;
		};
		x[PAD - k] = x[PAD + mirror(-k)]!;
		x[PAD + count - 1 + k] = x[PAD + mirror(count - 1 + k)]!;
	}
	// Index j of x holds absolute position i0 - PAD + j; even positions are low-pass samples.
	const first = (parity: number, from: number) =>
		from + (((i0 - PAD + from) & 1) === parity ? 0 : 1);
	const end = count + 2 * PAD;
	if (reversible) {
		for (let j = first(0, 1); j < end - 1; j += 2)
			x[j] = x[j]! - Math.floor((x[j - 1]! + x[j + 1]! + 2) / 4);
		for (let j = first(1, 2); j < end - 2; j += 2)
			x[j] = x[j]! + Math.floor((x[j - 1]! + x[j + 1]!) / 2);
		return;
	}
	for (let j = first(0, 0); j < end; j += 2) x[j] = x[j]! * K;
	for (let j = first(1, 0); j < end; j += 2) x[j] = x[j]! / K;
	const lift = (parity: number, margin: number, factor: number) => {
		for (let j = first(parity, margin); j < end - margin; j += 2)
			x[j] = x[j]! - factor * (x[j - 1]! + x[j + 1]!);
	};
	lift(0, 1, DELTA);
	lift(1, 2, GAMMA);
	lift(0, 3, BETA);
	lift(1, 4, ALPHA);
}

interface Plane {
	x0: number;
	y0: number;
	width: number;
	height: number;
	data: Float32Array;
}

/** The tile-component samples from its sub-band coefficients (`bands` per resolution). */
export function inverseDwt(
	component: TileComponent,
	coefficients: Float32Array[][],
	budget: Budget,
): Float32Array {
	const reversible = component.style.reversible;
	const base = component.resolutions[0]!;
	let low: Plane = {
		x0: base.x0,
		y0: base.y0,
		width: base.x1 - base.x0,
		height: base.y1 - base.y0,
		data: coefficients[0]![0]!,
	};
	for (let r = 1; r < component.resolutions.length; r++) {
		const res = component.resolutions[r]!;
		const width = res.x1 - res.x0;
		const height = res.y1 - res.y0;
		budget.spend(width * height);
		const data = new Float32Array(width * height);
		const [hl, lh, hh] = res.bands;
		const sources = [
			low,
			...[hl!, lh!, hh!].map((band, index) => ({
				x0: band.x0,
				y0: band.y0,
				width: band.x1 - band.x0,
				height: band.y1 - band.y0,
				data: coefficients[r]![index]!,
			})),
		];
		// 2D_INTERLEAVE: low-pass samples at even positions of the reference grid.
		for (let v = res.y0; v < res.y1; v++)
			for (let u = res.x0; u < res.x1; u++) {
				const source = sources[(u & 1) | ((v & 1) << 1)]!;
				const sx = Math.floor(u / 2) - source.x0;
				const sy = Math.floor(v / 2) - source.y0;
				data[(v - res.y0) * width + u - res.x0] =
					sx >= 0 && sy >= 0 && sx < source.width && sy < source.height
						? source.data[sy * source.width + sx]!
						: 0;
			}
		if (line.length < Math.max(width, height) + 2 * PAD)
			line = new Float64Array(Math.max(width, height) + 2 * PAD);
		if (width > 0)
			for (let y = 0; y < height; y++) {
				const row = y * width;
				for (let x = 0; x < width; x++) line[PAD + x] = data[row + x]!;
				filter(line, res.x0, width, reversible);
				for (let x = 0; x < width; x++) data[row + x] = line[PAD + x]!;
			}
		if (height > 0)
			for (let x = 0; x < width; x++) {
				for (let y = 0; y < height; y++) line[PAD + y] = data[y * width + x]!;
				filter(line, res.y0, height, reversible);
				for (let y = 0; y < height; y++) data[y * width + x] = line[PAD + y]!;
			}
		low = { x0: res.x0, y0: res.y0, width, height, data };
	}
	return low.data;
}
