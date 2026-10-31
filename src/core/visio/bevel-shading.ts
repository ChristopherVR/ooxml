/**
 * Bevel profiles, lighting and materials, DOM-free: the numbers the SVG bevel filter is built from.
 *
 * A bevel is drawn as a height field. Each of the 12 gallery bevels has a profile h(t): the surface
 * height at a fraction t of the bevel width from the outer edge (t = 0) to the face (t = 1). The
 * profiles are shaped by eye from Visio 16 raster exports of a 2 x 1.5 in rectangle with a 0.3 in
 * bevel (a column through the top edge, a row through the left edge), so they reproduce each
 * bevel's character (convex, concave, stepped, grooved) and not its exact curves; the corners are
 * rounded where Visio mitres them.
 *
 * `BevelLightingType` (1-15, 0 and anything above 15 read as 1) is the gallery's lighting. Each
 * entry holds the brightness of the face and of the four facets of the Angle bevel (a 45 degree
 * chamfer) relative to the unlit fill, measured from the same exports; `bevelLight` fits one
 * directional light and an ambient term to them. `BevelMaterialType` is read for the flat material
 * (7, no shading) and the three powder materials (9-11, which lighten the face and soften the
 * shading); the others, including the opaque white one (8), draw as the default material.
 * `BevelLightingAngle` changed nothing in the export, so it is not read.
 */
export interface BevelPoint {
	t: number;
	h: number;
}
type Anchors = readonly (readonly [number, number])[];

/** Points of a quarter circle from (t0, h0) rising to (t1, h1); convex means steep at the start. */
function arc(t0: number, h0: number, t1: number, h1: number, convex: boolean, steps = 6): Anchors {
	const result: [number, number][] = [];
	for (let i = 0; i <= steps; i++) {
		const u = i / steps;
		const k = convex ? Math.sqrt(1 - (1 - u) * (1 - u)) : 1 - Math.sqrt(1 - u * u);
		result.push([t0 + (t1 - t0) * u, h0 + (h1 - h0) * k]);
	}
	return result;
}
const smooth = (t0: number, h0: number, t1: number, h1: number, steps = 8): Anchors =>
	Array.from({ length: steps + 1 }, (_, i) => {
		const u = i / steps;
		return [t0 + (t1 - t0) * u, h0 + (h1 - h0) * u * u * (3 - 2 * u)] as const;
	});

const PROFILES: Record<number, Anchors> = {
	// Circle: a quarter circle, steep at the outer edge.
	1: arc(0, 0, 1, 1, true, 10),
	// Relaxed Inset: a rounded rim, then a recess that slopes back down to the face.
	2: [...arc(0, 0.2, 0.42, 1, true), [0.5, 1], ...smooth(0.5, 1, 1, 0.45)],
	// Cross: flat with one narrow groove.
	3: [
		[0, 1],
		[0.5, 1],
		[0.57, 0.3],
		[0.64, 1],
		[1, 1],
	],
	// Cool Slant: a short, steep convex slant on the outer edge.
	4: [...arc(0, 0, 0.35, 1, true), [1, 1]],
	// Angle: a straight chamfer (45 degrees when the bevel is as high as it is wide).
	5: [
		[0, 0],
		[1, 1],
	],
	// Soft Round: a dish; the rim is high and the surface sinks to the face.
	6: smooth(0, 1, 1, 0.5, 10),
	// Convex: a parabolic dome, steep outside and level at the face.
	7: Array.from({ length: 11 }, (_, i) => [i / 10, 1 - (1 - i / 10) ** 2] as const),
	// Slope: a gentle ramp that steepens towards the face.
	8: Array.from({ length: 11 }, (_, i) => [i / 10, (i / 10) ** 2] as const),
	// Divot: a rounded edge and a small dimple behind it.
	9: [...arc(0, 0, 0.3, 1, true), [0.34, 1], [0.4, 0.82], [0.46, 1], [1, 1]],
	// Riblet: a rounded edge and a thin rib near the face.
	10: [...arc(0, 0, 0.28, 0.85, true), [0.7, 0.85], [0.8, 1], [0.9, 0.85], [1, 0.85]],
	// Hard Edge: a very narrow, steep edge and a thin groove.
	11: [
		[0, 0],
		[0.12, 1],
		[0.66, 1],
		[0.72, 0.6],
		[0.78, 1],
		[1, 1],
	],
	// Art Deco: two steps with a groove between them.
	12: [
		[0, 0],
		[0.12, 0.45],
		[0.3, 0.45],
		[0.36, 0.3],
		[0.42, 0.45],
		[0.58, 0.45],
		[0.7, 1],
		[1, 1],
	],
};

/** The profile of a bevel type (1-12; others read as 5, a plain chamfer) as height points. */
export function bevelProfile(type: number): BevelPoint[] {
	return (PROFILES[type] ?? PROFILES[5]!).map(([t, h]) => ({ t, h }));
}
/** Height of a profile at `t` in 0-1 (linear between points, clamped). */
export function bevelHeightAt(points: readonly BevelPoint[], t: number): number {
	if (t <= points[0]!.t) return points[0]!.h;
	for (let i = 1; i < points.length; i++) {
		const [a, b] = [points[i - 1]!, points[i]!];
		if (t <= b.t) return b.t === a.t ? b.h : a.h + ((b.h - a.h) * (t - a.t)) / (b.t - a.t);
	}
	return points[points.length - 1]!.h;
}

/** The inverse normal distribution for 0.5 < p < 1 (Winitzki's erf approximation, refined). */
function normalQuantile(p: number): number {
	const x = 2 * p - 1;
	const a = 0.147;
	const ln = Math.log(1 - x * x);
	const part = 2 / (Math.PI * a) + ln / 2;
	const erfInv = Math.sqrt(Math.sqrt(part * part - ln / a) - part);
	return Math.SQRT2 * erfInv;
}
/** The standard deviation of the blur the height field is read through, in bevel widths. */
export const BEVEL_BLUR_WIDTHS = 1 / 2.2;

export interface BevelHeightTable {
	/** Heights scaled to 0-1, for an `feComponentTransfer` table over the blurred alpha. */
	values: number[];
	/** Rise of the profile (highest minus lowest) in units of the bevel height. */
	rise: number;
}
/**
 * The table that turns a blurred alpha (an edge blurred with `BEVEL_BLUR_WIDTHS` bevel widths reads
 * Phi(d / sigma) at depth d) back into the bevel's height at that depth, so a height-field lighting
 * filter sees the profile instead of the blur's own S-curve. Depths beyond the bevel width are the
 * face. Outside the edge (alpha at or below one half) the height is the profile's start.
 */
export function bevelHeightTable(type: number, entries = 128): BevelHeightTable {
	const points = bevelProfile(type);
	const heights: number[] = [];
	for (let i = 0; i < entries; i++) {
		const alpha = i / (entries - 1);
		const depth =
			alpha <= 0.5 ? 0 : Math.min(1, BEVEL_BLUR_WIDTHS * normalQuantile(Math.min(alpha, 0.9999)));
		heights.push(bevelHeightAt(points, depth));
	}
	const low = Math.min(...heights);
	const high = Math.max(...heights);
	const range = high - low || 1;
	return {
		values: heights.map((h) => Math.round(((h - low) / range) * 1e4) / 1e4),
		rise: range,
	};
}

/** Face, top, right, bottom, left brightness of the Angle bevel against the unlit fill. */
const LIGHTING_GAINS: Record<number, readonly [number, number, number, number, number]> = {
	1: [1.04, 1.09, 0.6, 0.51, 0.87],
	2: [0.96, 1.16, 0.69, 1.07, 0.98],
	3: [0.96, 1.09, 0.96, 0.7, 0.82],
	4: [0.84, 0.84, 0.28, 0.88, 1.55],
	5: [0.86, 0.95, 0.34, 0.8, 2.28],
	6: [1, 1, 2.28, 1, 2.28],
	7: [0.85, 0.37, 1.12, 1.21, 1.46],
	8: [0.79, 0.38, 0.87, 1.1, 1.34],
	9: [0.7, 0.38, 0.53, 0.89, 0.96],
	10: [0.99, 0.64, 0.55, 1.26, 1.39],
	11: [0.75, 0.38, 1.53, 1.01, 1.25],
	12: [1, 1.29, 1.01, 1, 1],
	13: [0.98, 0.66, 0.58, 1.0, 0.96],
	14: [1, 1, 1.59, 1, 2.28],
	15: [1, 1.14, 1.5, 1.14, 1.85],
};
export const BEVEL_LIGHTING_TYPES = Object.keys(LIGHTING_GAINS).length;
/** A lighting gallery type; 0 (the default) and out-of-range values are the first, three point. */
export const normalBevelLighting = (type: number): number =>
	Number.isInteger(type) && type >= 1 && type <= BEVEL_LIGHTING_TYPES ? type : 1;

export interface BevelLight {
	/** Direction to the light on the page (degrees clockwise from the right, y down). */
	azimuth: number;
	/** Height of the light above the page plane in degrees. */
	elevation: number;
	/** Brightness of a surface facing the light head on, against the unlit fill. */
	diffuse: number;
	/** Brightness of a surface the light does not reach. */
	ambient: number;
	/** Brightness of the flat face. */
	face: number;
}
const FACET = Math.SQRT1_2;
/**
 * One directional light and an ambient term fitted to the measured facet brightness of the Angle
 * bevel: brightness = ambient + diffuse * (n . L), with the facet normals tilted 45 degrees.
 */
export function bevelLight(lighting: number): BevelLight {
	const [face, top, right, bottom, left] = LIGHTING_GAINS[normalBevelLighting(lighting)]!;
	const lx = (right - left) / (2 * FACET);
	const ly = (bottom - top) / (2 * FACET);
	const lz = (face - (top + right + bottom + left) / 4) / (1 - FACET);
	const diffuse = Math.hypot(lx, ly, lz) || 1;
	return {
		azimuth: (Math.atan2(ly, lx) * 180) / Math.PI,
		elevation: (Math.asin(Math.max(-1, Math.min(1, lz / diffuse))) * 180) / Math.PI,
		diffuse,
		ambient: face - lz,
		face,
	};
}

export interface BevelMaterial {
	/** Multiplier of the shading's departure from the face. */
	contrast: number;
	/** Fraction of white mixed into the face (the powder materials). */
	lift: number;
}
/** Material 7 is flat (no shading) and 9-11 are the powder materials; every other value is plain. */
export function bevelMaterial(type: number): BevelMaterial {
	switch (type) {
		case 7:
			return { contrast: 0, lift: 0 };
		case 9:
			return { contrast: 0.7, lift: 0.2 };
		case 10:
			return { contrast: 0.45, lift: 0.5 };
		case 11:
			return { contrast: 0.45, lift: 0.8 };
		default:
			return { contrast: 1, lift: 0 };
	}
}
/** Materials the renderer cannot draw (8 is an opaque white one) and that draw as the default. */
export const isUndrawnBevelMaterial = (type: number): boolean => type === 8;
