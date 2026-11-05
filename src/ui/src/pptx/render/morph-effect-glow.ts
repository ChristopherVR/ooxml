/**
 * The `filter` that cross-dissolves a morphing shape's two glows the way
 * PowerPoint does.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object,
 * 2 s) on a rectangle with a 20 pt yellow glow growing from 200 x 120 pt to
 * 440 x 240 pt: the halo beyond the body's right edge is fitted within 0.008
 * (0-1 of full glow) at every frame by `(1 - p) G(x / k0) + p G(x / k1)`, the
 * outgoing and the incoming glow each riding the box (`k = W(p) / W`) and
 * dissolving on the morph's eased progress, and only within 0.017 by the best
 * single glow of any width. So PowerPoint dissolves two glows, as it does two
 * shadows, rather than widening one.
 *
 * A glow is a `drop-shadow()`, and a filter list cannot sum two of them: the
 * second one shadows the first one's output, glow and all. Painting the
 * NARROWER glow first and the wider one over the result keeps that error
 * small, and the two alphas are fitted per keyframe (Gauss-Newton on a 1-D
 * profile across a straight edge, `drop-shadow` being a Gaussian of the
 * silhouette) so the chain lands on the dissolve; in that model the fit is
 * within 0.002 alpha of the dissolve where the plain fades are 0.005-0.02 off
 * and one widening glow 0.007-0.01 off.
 *
 * @module render/morph-effect-glow
 */

/** One end's glow: its blur radius in the half's local px and its colour alpha. */
export interface GlowEnd {
	readonly radius: number;
	readonly alpha: number;
}

const SAMPLES = 64;

/** `erfc` (Abramowitz and Stegun 7.1.26, error under 1.5e-7). */
function erfc(x: number): number {
	const z = Math.abs(x);
	const t = 1 / (1 + 0.3275911 * z);
	const poly =
		t *
		(0.254829592 + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
	const value = poly * Math.exp(-z * z);
	return x >= 0 ? value : 2 - value;
}

/** A blurred straight edge's alpha at distance `x` outside it. */
function edge(x: number, sigma: number): number {
	return sigma <= 0 ? (x < 0 ? 1 : 0) : 0.5 * erfc(x / (sigma * Math.SQRT2));
}

/** The 1-D profile both models are compared on. */
interface Profile {
	readonly xs: number[];
	readonly step: number;
	readonly narrow: number;
	readonly wide: number;
	/** `edge(x, narrow)` per sample. */
	readonly sn: number[];
	/** `edge(x, wide)` per sample. */
	readonly sw: number[];
	/** The blur of the narrow glow's outside part by the wide one, per unit alpha. */
	readonly glowOfGlow: number[];
}

function profile(narrow: number, wide: number): Profile {
	// `drop-shadow`'s length is a radius; the Gaussian's deviation is half of it.
	const sn = Math.max(narrow / 2, 1e-3);
	const sw = Math.max(wide / 2, sn);
	const span = 3.5 * (sw + sn);
	const step = span / (SAMPLES - 1);
	const xs = Array.from({ length: SAMPLES }, (_, i) => i * step);
	const narrowEdge = xs.map((x) => edge(x, sn));
	const glowOfGlow = xs.map((x) => {
		let sum = 0;
		for (let j = 0; j < SAMPLES; j++) {
			const d = x - (xs[j] ?? 0);
			sum += (narrowEdge[j] ?? 0) * Math.exp(-(d * d) / (2 * sw * sw));
		}
		return (sum * step) / (sw * Math.sqrt(2 * Math.PI));
	});
	return {
		xs,
		step,
		narrow: sn,
		wide: sw,
		sn: narrowEdge,
		sw: xs.map((x) => edge(x, sw)),
		glowOfGlow,
	};
}

/** The chain `drop-shadow(narrow, a1) drop-shadow(wide, a2)` outside the edge. */
function chain(prof: Profile, a1: number, a2: number): number[] {
	return prof.xs.map((_, i) => {
		const first = a1 * (prof.sn[i] ?? 0);
		const blurred = (prof.sw[i] ?? 0) + a1 * (prof.glowOfGlow[i] ?? 0);
		return first + (1 - first) * a2 * Math.min(1, blurred);
	});
}

function clampUnit(value: number): number {
	return Math.min(1, Math.max(0, value));
}

/**
 * The two alphas of the chain (narrow glow first, wide glow second) that best
 * paint `(1 - p) A + p B` at progress `p`, where `a` is the outgoing glow and
 * `b` the incoming one, both in the same frame.
 */
export function fitGlowChain(a: GlowEnd, b: GlowEnd, p: number): [number, number] {
	const aNarrow = a.radius <= b.radius;
	const prof = profile(Math.min(a.radius, b.radius), Math.max(a.radius, b.radius));
	const fade = (end: GlowEnd, sigmaEdge: number[]) => sigmaEdge.map((s) => end.alpha * s);
	const [narrowEnd, wideEnd] = aNarrow ? [a, b] : [b, a];
	const [narrowWeight, wideWeight] = aNarrow ? [1 - p, p] : [p, 1 - p];
	const ta = fade(narrowEnd, prof.sn);
	const tb = fade(wideEnd, prof.sw);
	const target = ta.map((v, i) => narrowWeight * v + wideWeight * (tb[i] ?? 0));
	const residual = (x1: number, x2: number) =>
		chain(prof, x1, x2).map((v, i) => v - (target[i] ?? 0));
	let x1 = narrowWeight * narrowEnd.alpha;
	let x2 = wideWeight * wideEnd.alpha;
	const h = 1e-4;
	for (let iter = 0; iter < 8; iter++) {
		const r = residual(x1, x2);
		const r1 = residual(x1 + h, x2);
		const r2 = residual(x1, x2 + h);
		let j11 = 0;
		let j12 = 0;
		let j22 = 0;
		let g1 = 0;
		let g2 = 0;
		r.forEach((value, i) => {
			const d1 = ((r1[i] ?? 0) - value) / h;
			const d2 = ((r2[i] ?? 0) - value) / h;
			j11 += d1 * d1;
			j12 += d1 * d2;
			j22 += d2 * d2;
			g1 += d1 * value;
			g2 += d2 * value;
		});
		const det = j11 * j22 - j12 * j12;
		if (Math.abs(det) < 1e-12) {
			break;
		}
		const s1 = (j22 * g1 - j12 * g2) / det;
		const s2 = (j11 * g2 - j12 * g1) / det;
		x1 = clampUnit(x1 - s1);
		x2 = clampUnit(x2 - s2);
		if (Math.abs(s1) + Math.abs(s2) < 1e-5) {
			break;
		}
	}
	return [x1, x2];
}

/**
 * The incoming glow's alpha under a ghost that paints the outgoing glow at
 * `1 - p` OVER it, so the two composite to `(1 - p) A + p B`. Least squares
 * on the same profile.
 */
export function fitGlowUnderGhost(a: GlowEnd, b: GlowEnd, p: number): number {
	const prof = profile(Math.min(a.radius, b.radius), Math.max(a.radius, b.radius));
	const sa = a.radius <= b.radius ? prof.sn : prof.sw;
	const sb = a.radius <= b.radius ? prof.sw : prof.sn;
	let num = 0;
	let den = 0;
	sb.forEach((s, i) => {
		const ghost = (1 - p) * a.alpha * (sa[i] ?? 0);
		const basis = s * (1 - ghost);
		num += basis * p * b.alpha * s;
		den += basis * basis;
	});
	return den > 0 ? clampUnit(num / den) : p * b.alpha;
}
