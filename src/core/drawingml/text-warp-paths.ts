import { archEllipsePath, circleEllipsePath } from './text-warp-arch';

export const clamp4 = (n: number): number => Math.max(0, Math.min(n, 4));

// ── Priority 1 ──────────────────────────────────────────────────────────

/** Concentric upward arcs. t=0 is the tallest arch. */
export function archUpPath(w: number, h: number, _t: number, adj?: number): string {
	return archEllipsePath(w, h, 0, adj, 'up');
}

/** Concentric downward arcs. t=1 is the deepest. */
export function archDownPath(w: number, h: number, _t: number, adj?: number): string {
	return archEllipsePath(w, h, 0, adj, 'down');
}

/** Full ellipse once around clockwise, from the tip the start angle picks (see `circleEllipsePath`). */
export function circlePath(w: number, h: number, _t: number, adj?: number): string {
	return circleEllipsePath(w, h, 0, adj);
}

/** Single sine-wave via cubic Bézier. */
export function wave1Path(w: number, h: number, t: number, adj?: number, adj2?: number): string {
	const yMid = h * (0.25 + t * 0.5);
	const adjFactor = adj !== undefined ? adj / 12500 : 1;
	const amp = h * 0.2 * clamp4(adjFactor);
	const hShift = adj2 !== undefined ? (adj2 / 100000) * w * 0.3 : 0;
	const cp1x = w / 3 + hShift;
	const cp2x = (2 * w) / 3 + hShift;
	return `M 0,${yMid} C ${cp1x},${yMid - amp} ${cp2x},${yMid + amp} ${w},${yMid}`;
}

/**
 * Inflate: the baseline bows up in the middle (balloon shape).
 *
 * `intensity` sets a BASE amplitude that stays present at `t=0.5` (a single
 * paragraph, the common WordArt case) plus a smaller per-line `extra` term so
 * a multi-line block still varies top-to-bottom. The original formula used
 * `(1 - 2 * t)` alone, which is exactly zero at `t=0.5`: single-line Inflate
 * text rendered perfectly flat, indistinguishable from no warp at all.
 */
export function inflatePath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const intensity = adj !== undefined ? clamp4(adj / 18750) : 1;
	const base = h * 0.22 * intensity;
	const extra = h * 0.1 * intensity * (0.5 - t);
	const bulge = base + extra;
	return `M 0,${yBase} Q ${w / 2},${yBase - bulge} ${w},${yBase}`;
}

/**
 * Deflate: the baseline dips down in the middle (pinched shape), the
 * opposite curvature from {@link inflatePath}. Same base+extra fix for the
 * `t=0.5` single-line degeneracy the original `(2 * t - 1)` formula had.
 */
export function deflatePath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const intensity = adj !== undefined ? clamp4(adj / 18750) : 1;
	const base = h * 0.18 * intensity;
	const extra = h * 0.08 * intensity * (t - 0.5);
	const dip = base + extra;
	return `M 0,${yBase} Q ${w / 2},${yBase + dip} ${w},${yBase}`;
}

/** Gentle upward curve. */
export function curveUpPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.35 + t * 0.55);
	const adjFactor = adj !== undefined ? adj / 45977 : 1;
	const curve = h * 0.4 * (1 - t * 0.3) * clamp4(adjFactor);
	return `M 0,${yBase} Q ${w / 2},${yBase - curve} ${w},${yBase}`;
}

/** Gentle downward curve. */
export function curveDownPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.1 + t * 0.55);
	const adjFactor = adj !== undefined ? adj / 45977 : 1;
	const curve = h * 0.4 * (1 - (1 - t) * 0.3) * clamp4(adjFactor);
	return `M 0,${yBase} Q ${w / 2},${yBase + curve} ${w},${yBase}`;
}

// ── Priority 2 ──────────────────────────────────────────────────────────

/** Inverted single wave (phase-shifted wave1). */
export function wave2Path(w: number, h: number, t: number, adj?: number, adj2?: number): string {
	const yMid = h * (0.25 + t * 0.5);
	const adjFactor = adj !== undefined ? adj / 12500 : 1;
	const amp = h * 0.2 * clamp4(adjFactor);
	const hShift = adj2 !== undefined ? (adj2 / 100000) * w * 0.3 : 0;
	const cp1x = w / 3 + hShift;
	const cp2x = (2 * w) / 3 + hShift;
	return `M 0,${yMid} C ${cp1x},${yMid + amp} ${cp2x},${yMid - amp} ${w},${yMid}`;
}

/** Double wave: two full cycles. */
export function wave4Path(w: number, h: number, t: number, adj?: number, adj2?: number): string {
	const yMid = h * (0.25 + t * 0.5);
	const adjFactor = adj !== undefined ? adj / 12500 : 1;
	const amp = h * 0.15 * clamp4(adjFactor);
	const hShift = adj2 !== undefined ? (adj2 / 100000) * w * 0.15 : 0;
	const q = w / 4;
	return (
		`M 0,${yMid} ` +
		`C ${q + hShift},${yMid - amp} ${2 * q + hShift},${yMid + amp} ${w / 2},${yMid} ` +
		`C ${w / 2 + q + hShift},${yMid - amp} ${w - q + hShift},${yMid + amp} ${w},${yMid}`
	);
}

/** Double wave with alternating rhythm. */
export function doubleWave1Path(
	w: number,
	h: number,
	t: number,
	adj?: number,
	adj2?: number,
): string {
	const yMid = h * (0.25 + t * 0.5);
	const adjFactor = adj !== undefined ? adj / 6250 : 1;
	const amp = h * 0.18 * clamp4(adjFactor);
	const hShift = adj2 !== undefined ? (adj2 / 100000) * w * 0.15 : 0;
	const q = w / 4;
	return (
		`M 0,${yMid} ` +
		`C ${q + hShift},${yMid - amp} ${2 * q + hShift},${yMid + amp} ${w / 2},${yMid} ` +
		`C ${w / 2 + q + hShift},${yMid + amp} ${w - q + hShift},${yMid - amp} ${w},${yMid}`
	);
}

/** Cylindrical text: upward. */
export function canUpPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const curvature = clamp4(adjFactor);
	const archH = h * (0.35 - t * 0.25) * curvature;
	if (archH < 1) {
		return `M 0,${h} L ${w},${h}`;
	}
	return `M 0,${h} A ${w / 2},${archH} 0 0,1 ${w},${h}`;
}

/** Cylindrical text: downward. */
export function canDownPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const curvature = clamp4(adjFactor);
	const archH = h * (0.1 + t * 0.25) * curvature;
	if (archH < 1) {
		return `M 0,0 L ${w},0`;
	}
	return `M 0,0 A ${w / 2},${archH} 0 0,0 ${w},0`;
}

/** Button: a single line lies on the same arc as Arch Up (and turns with the start angle alike). */
export function buttonPath(w: number, h: number, _t: number, adj?: number): string {
	return archEllipsePath(w, h, 0, adj, 'up');
}

/** Ring inside: concentric ellipses scaled inward. */
export function ringInsidePath(w: number, h: number, t: number, adj?: number): string {
	const cx = w / 2;
	const cy = h / 2;
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const thickness = 0.35 * clamp4(adjFactor);
	const scale = 0.7 - t * thickness;
	const rx = Math.max(1, (w / 2) * scale);
	const ry = Math.max(1, (h / 2) * scale);
	return (
		`M ${cx},${cy - ry} ` +
		`A ${rx},${ry} 0 1,1 ${cx},${cy + ry} ` +
		`A ${rx},${ry} 0 1,1 ${cx},${cy - ry}`
	);
}

/** Ring outside: concentric ellipses scaled outward. */
export function ringOutsidePath(w: number, h: number, t: number, adj?: number): string {
	const cx = w / 2;
	const cy = h / 2;
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const thickness = 0.35 * clamp4(adjFactor);
	const scale = 1 - t * thickness;
	const rx = Math.max(1, (w / 2) * scale);
	const ry = Math.max(1, (h / 2) * scale);
	return (
		`M ${cx},${cy - ry} ` +
		`A ${rx},${ry} 0 1,1 ${cx},${cy + ry} ` +
		`A ${rx},${ry} 0 1,1 ${cx},${cy - ry}`
	);
}
