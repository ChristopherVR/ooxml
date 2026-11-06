/**
 * Geometry of `circularArrow`, `leftCircularArrow` and `leftRightCircularArrow`.
 *
 * The three presets share one construction (ECMA-376 `presetShapeDefinitions`):
 * a band of thickness `th` centred on an ellipse of radii `rw3` x `rh3`, whose
 * head is a triangle `thh` wide on each side of the band centre. The head base
 * centre `H` lies on the centre line at the band's end angle `enAng`, the tip
 * `A` at `enAng + aAng` and the base points `G` (outer) and `B` (inner) are
 * `thh` either side of `H` along the tip's radial direction. The band's outer
 * and inner arcs stop where the base line crosses the outer and inner ellipse.
 *
 * The head angle is limited to `maxAng`, the angle from `H` to the point where
 * the tangent from `H` touches the inner circle, so the tip never turns back
 * into the band.
 */

import type { PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

/** Reduce `x` (60000ths, between -360 and 720 degrees) to the range 0..360. */
function norm(name: string, x: string): string[] {
	return [
		`${name}a +- ${x} 21600000 0`,
		`${name}b ?: ${x} ${x} ${name}a`,
		`${name}c +- ${name}b 0 21600000`,
		`${name} ?: ${name}c ${name}c ${name}b`,
	];
}

/** A point on the ellipse `rw` x `rh` seen along the visual angle `ang`. */
function onEllipse(name: string, rw: string, rh: string, ang: string): string[] {
	return [
		`${name}wt sin ${rw} ${ang}`,
		`${name}ht cos ${rh} ${ang}`,
		`${name}dx cat2 ${rw} ${name}ht ${name}wt`,
		`${name}dy sat2 ${rh} ${name}ht ${name}wt`,
		`${name}x +- hc ${name}dx 0`,
		`${name}y +- vc ${name}dy 0`,
	];
}

/**
 * One arrow head: base centre at `baseAng`, tip at `tipAng`. Gives the tip
 * (`<n>A`), the base points (`<n>G` outer, `<n>B` inner) and where the base line
 * leaves the outer (`<n>F`) and the inner (`<n>D`) ellipse, with their angles.
 */
function head(n: string, baseAng: string, tipAng: string): string[] {
	const out: string[] = [
		...onEllipse(`${n}H`, 'rw3', 'rh3', baseAng),
		...onEllipse(`${n}A`, 'rw3', 'rh3', tipAng),
		`${n}Gdx cos thh ${tipAng}`,
		`${n}Gdy sin thh ${tipAng}`,
		`${n}Gx +- ${n}Hx ${n}Gdx 0`,
		`${n}Gy +- ${n}Hy ${n}Gdy 0`,
		`${n}Bx +- ${n}Hx 0 ${n}Gdx`,
		`${n}By +- ${n}Hy 0 ${n}Gdy`,
		`${n}sx1 +- ${n}Bx 0 hc`,
		`${n}sy1 +- ${n}By 0 vc`,
		`${n}sx2 +- ${n}Gx 0 hc`,
		`${n}sy2 +- ${n}Gy 0 vc`,
	];
	// Cross the line B -> G with the outer (rw1, rh1) and inner (rw2, rh2) ellipse:
	// scale to a circle of radius r, take the root further along B -> G, scale back.
	for (const [k, rw, rh] of [
		['F', 'rw1', 'rh1'],
		['D', 'rw2', 'rh2'],
	] as const) {
		const m = `${n}${k}`;
		out.push(
			`${m}r min ${rw} ${rh}`,
			`${m}x1 */ ${n}sx1 ${m}r ${rw}`,
			`${m}y1 */ ${n}sy1 ${m}r ${rh}`,
			`${m}x2 */ ${n}sx2 ${m}r ${rw}`,
			`${m}y2 */ ${n}sy2 ${m}r ${rh}`,
			`${m}ex +- ${m}x2 0 ${m}x1`,
			`${m}ey +- ${m}y2 0 ${m}y1`,
			`${m}len mod ${m}ex ${m}ey 0`,
			`${m}pd1 */ ${m}x1 ${m}ex ${m}len`,
			`${m}pd2 */ ${m}y1 ${m}ey ${m}len`,
			`${m}pd +- ${m}pd1 ${m}pd2 0`,
			`${m}pd2s */ ${m}pd ${m}pd 1`,
			`${m}p1a */ ${m}x1 ${m}x1 1`,
			`${m}p1b */ ${m}y1 ${m}y1 1`,
			`${m}rsq */ ${m}r ${m}r 1`,
			`${m}d0 +- ${m}rsq 0 ${m}p1a`,
			`${m}d1 +- ${m}d0 0 ${m}p1b`,
			`${m}d2 +- ${m}d1 ${m}pd2s 0`,
			`${m}sq sqrt ${m}d2`,
			`${m}t +- ${m}sq 0 ${m}pd`,
			`${m}ux */ ${m}t ${m}ex ${m}len`,
			`${m}uy */ ${m}t ${m}ey ${m}len`,
			`${m}sx +- ${m}x1 ${m}ux 0`,
			`${m}sy +- ${m}y1 ${m}uy 0`,
			`${m}dx */ ${m}sx ${rw} ${m}r`,
			`${m}dy */ ${m}sy ${rh} ${m}r`,
			`${m}X +- hc ${m}dx 0`,
			`${m}Y +- vc ${m}dy 0`,
			`${m}ang at2 ${m}dx ${m}dy`,
		);
	}
	return out;
}

/** Thickness, radii and the head limit shared by all three presets. */
const BASE = [
	'a5 pin 0 adj5 25000',
	'maxAdj1 */ a5 2 1',
	'a1 pin 0 adj1 maxAdj1',
	'enAng pin 1 adj3 21599999',
	'stAng pin 0 adj4 21599999',
	'th */ ss a1 100000',
	'thh */ ss a5 100000',
	'th2 */ th 1 2',
	'rw1 +- wd2 th2 thh',
	'rh1 +- hd2 th2 thh',
	'rw2 +- rw1 0 th',
	'rh2 +- rh1 0 th',
	'rw3 +- rw2 th2 0',
	'rh3 +- rh2 th2 0',
	...onEllipse('H', 'rw3', 'rh3', 'enAng'),
	'rI min rw2 rh2',
	'u1 */ Hdx Hdx 1',
	'u2 */ Hdy Hdy 1',
	'u3 */ rI rI 1',
	'u4 +- u1 u2 u3',
	'tq sqrt u4',
	'tAng at2 rI tq',
	'hAng at2 Hdx Hdy',
];

/** Outline order of one head: `F` -> `G` -> tip -> `B` -> `D`. */
const headPath = (n: string) => `L ${n}Gx ${n}Gy L ${n}Ax ${n}Ay L ${n}Bx ${n}By L ${n}DX ${n}DY`;

const circularArrow = exactDef(
	'circularArrow',
	{ adj1: 12500, adj2: 1142319, adj3: 20457681, adj4: 10800000, adj5: 12500 },
	[
		...BASE,
		'mx1 +- hAng tAng enAng',
		...norm('maxAng', 'mx1'),
		'aAng pin 0 adj2 maxAng',
		'ptAng +- enAng aAng 0',
		...head('e', 'enAng', 'ptAng'),
		...onEllipse('so', 'rw1', 'rh1', 'stAng'),
		...onEllipse('si', 'rw2', 'rh2', 'stAng'),
		'swO0 +- eFang 0 stAng',
		...norm('swO', 'swO0'),
		'swI0 +- eDang 0 stAng',
		...norm('swI1', 'swI0'),
		'swI +- 0 0 swI1',
	],
	'l t r b',
	[p(`M sox soy A rw1 rh1 stAng swO ${headPath('e')} A rw2 rh2 eDang swI Z`)],
);

const leftCircularArrow = exactDef(
	'leftCircularArrow',
	{ adj1: 12500, adj2: -1142319, adj3: 1142319, adj4: 10800000, adj5: 12500 },
	[
		...BASE,
		'mx1 +- enAng tAng hAng',
		...norm('mx', 'mx1'),
		'minAng +- 0 0 mx',
		'aAng pin minAng adj2 0',
		'ptAng +- enAng aAng 0',
		...head('e', 'enAng', 'ptAng'),
		...onEllipse('so', 'rw1', 'rh1', 'stAng'),
		...onEllipse('si', 'rw2', 'rh2', 'stAng'),
		'swO0 +- eFang 0 stAng',
		...norm('swO1', 'swO0'),
		'swO +- swO1 0 21600000',
		'swI0 +- stAng 0 eDang',
		...norm('swI', 'swI0'),
	],
	'l t r b',
	[p(`M sox soy A rw1 rh1 stAng swO ${headPath('e')} A rw2 rh2 eDang swI Z`)],
);

/** The second head of `leftRightCircularArrow`: base at `stAng`, tip behind it. */
const leftRightCircularArrow = exactDef(
	'leftRightCircularArrow',
	{ adj1: 12500, adj2: 1142319, adj3: 20457681, adj4: 11942319, adj5: 12500 },
	[
		...BASE,
		'mx1 +- hAng tAng enAng',
		...norm('maxAng', 'mx1'),
		'aAng pin 0 adj2 maxAng',
		'ptAng +- enAng aAng 0',
		'lptAng +- stAng 0 aAng',
		...head('e', 'enAng', 'ptAng'),
		...head('s', 'stAng', 'lptAng'),
		'swO0 +- eFang 0 sFang',
		...norm('swO', 'swO0'),
		'swI0 +- eDang 0 sDang',
		...norm('swI1', 'swI0'),
		'swI +- 0 0 swI1',
	],
	'l t r b',
	[
		p(
			`M sFX sFY A rw1 rh1 sFang swO ${headPath('e')} A rw2 rh2 eDang swI L sBx sBy L sAx sAy L sGx sGy Z`,
		),
	],
);

export const EXACT_CIRCULAR_ARROW_PRESET_DEFINITIONS: Record<
	string,
	PresetShapeGeometryDefinition
> = {
	circularArrow,
	leftCircularArrow,
	leftRightCircularArrow,
};
