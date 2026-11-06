/**
 * Geometry of the twelve `actionButton*` presets.
 *
 * Measured on PowerPoint 16 (150 x 90 pt, mid-grey fill, shaded): the button is
 * a flat rectangle (no bevel), and the glyph is drawn on it as filled sub-paths,
 * `darken` for the dark parts (arrows, house roof, speaker, camera, the circle of
 * the information button), `darkenLess` for the house body and the page, and
 * `lighten` for the white of the "i". Outlines of every sub-path are stroked
 * with the shape's line.
 *
 * Every glyph is placed relative to the centre `(hc, vc)` and scaled by `ss`;
 * the numbers below are points on a 90 pt `ss`, written once and turned into
 * guides here. The ECMA glyphs are polygons built on `dx2 = ss * 3 / 8`
 * (33.75); curved ones (the question mark) are a polyline through the measured
 * outline, not the specification's Bezier chain.
 */

import type { PresetPath, PresetShapeGeometryDefinition } from './preset-shape-definitions-table';
import { exactDef, p } from './preset-shape-exact-dsl';

type Pt = readonly [number, number];

/** Builds guides for coordinates given in points of a 90 pt `ss`. */
class Glyph {
	private readonly lines: string[] = [];
	private readonly names = new Map<string, string>();

	private coord(axis: 'x' | 'y', value: number): string {
		const key = `${axis}${value}`;
		const known = this.names.get(key);
		if (known) {
			return known;
		}
		const id = this.names.size;
		const name = `${axis}g${id}`;
		this.lines.push(
			`k${id} */ ss ${Math.round(value * 1000)} 90000`,
			`${name} +- ${axis === 'x' ? 'hc' : 'vc'} k${id} 0`,
		);
		this.names.set(key, name);
		return name;
	}

	/** A guide holding `value` points of a 90 pt `ss` (a radius, a length). */
	len(value: number): string {
		const key = `len${value}`;
		const known = this.names.get(key);
		if (known) {
			return known;
		}
		const name = `lg${this.names.size}`;
		this.lines.push(`${name} */ ss ${Math.round(value * 1000)} 90000`);
		this.names.set(key, name);
		return name;
	}

	x(v: number): string {
		return this.coord('x', v);
	}

	y(v: number): string {
		return this.coord('y', v);
	}

	pt(point: Pt): string {
		return `${this.x(point[0])} ${this.y(point[1])}`;
	}

	/** `M p0 L p1 ... Z` through the points. */
	poly(points: readonly Pt[], close = true): string {
		const [first, ...rest] = points;
		return `M ${this.pt(first!)}${rest.map((q) => ` L ${this.pt(q)}`).join('')}${close ? ' Z' : ''}`;
	}

	/** A full circle of radius `r` centred on `(cx, cy)`. */
	circle(cx: number, cy: number, r: number): string {
		return `M ${this.x(cx + r)} ${this.y(cy)} A ${this.len(r)} ${this.len(r)} 0 21600000 Z`;
	}

	guides(): string[] {
		return this.lines;
	}
}

const FACE = 'M l t L r t L r b L l b Z';

function build(name: string, glyph: (g: Glyph) => PresetPath[]): PresetShapeGeometryDefinition {
	const g = new Glyph();
	const paths = glyph(g);
	return exactDef(name, undefined, g.guides(), 'l t r b', [p(FACE), ...paths]);
}

const D = 33.75;
const darken = (src: string): PresetPath => p(src, { fill: 'darken' });
const darkenLess = (src: string): PresetPath => p(src, { fill: 'darkenLess' });
const lighten = (src: string): PresetPath => p(src, { fill: 'lighten' });

const triangle = (g: Glyph, dir: 1 | -1): string =>
	g.poly([
		[-D * dir, -D],
		[D * dir, 0],
		[-D * dir, D],
	]);

// the question mark outline, measured (points of a 90 pt ss)
const QUESTION: Pt[] = [
	[-4.8, -33.7],
	[-6.3, -31.5],
	[-7.8, -31.5],
	[-12.3, -28.5],
	[-14.5, -28.5],
	[-15.3, -27.7],
	[-15.3, -25.5],
	[-18.3, -21],
	[-19, -17.2],
	[-20.5, -16.5],
	[-20.5, -12.7],
	[-19, -12.7],
	[-18.3, -14.2],
	[-10.8, -14.2],
	[-10, -13.5],
	[-8.5, -13.5],
	[-9.3, -17.2],
	[-7.8, -19.5],
	[-4, -22.5],
	[-3.3, -21.7],
	[-1.8, -21.7],
	[-1, -23.2],
	[0.5, -23.2],
	[1.2, -21.7],
	[2.7, -21.7],
	[3.5, -22.5],
	[8, -18.7],
	[8.7, -17.2],
	[8, -16.5],
	[8, -15],
	[9.5, -14.2],
	[9.5, -13.5],
	[8, -12.7],
	[8.7, -9.7],
	[5, -6],
	[2, -6],
	[-2.5, -2.2],
	[-3.3, 0.8],
	[-5.5, 2.3],
	[-4.8, 6.8],
	[-6.3, 7.5],
	[-6.3, 18],
	[-4.8, 18],
	[-4, 17.3],
	[4.2, 17.3],
	[4.2, 6.8],
	[6.5, 5.3],
	[6.5, 3.8],
	[8, 3],
	[8, 1.5],
	[8.7, 0.8],
	[13.2, 0],
	[17, -4.5],
	[18.5, -4.5],
	[18.5, -6],
	[17.7, -6.7],
	[18.5, -8.2],
	[18.5, -15],
	[20, -15.7],
	[20, -18],
	[17.7, -19.5],
	[17.7, -21],
	[15.5, -24.7],
	[15.5, -27],
	[14, -27],
	[13.2, -27.7],
	[13.2, -29.2],
	[9.5, -30],
	[8.7, -30.7],
	[8.7, -32.2],
	[5.7, -31.5],
	[4.2, -33.7],
	[2.7, -33.7],
	[2, -32.2],
	[-2.5, -32.2],
	[-3.3, -33.7],
];

const MOVIE: Pt[] = [
	[-33.75, -16.5],
	[-33.75, -3],
	[-26, -3.7],
	[-26, 15.8],
	[19, 15.8],
	[19, 10.5],
	[20.5, 8.3],
	[25.7, 8.3],
	[30.2, 12],
	[33.75, 12],
	[33.75, -12],
	[29.5, -12.7],
	[27.2, -9],
	[20.5, -9],
	[19, -10.5],
	[18.2, -15.7],
	[-27.5, -15],
	[-29, -16.5],
];

const buttons: Record<string, PresetShapeGeometryDefinition> = {
	actionButtonBlank: build('actionButtonBlank', () => []),
	actionButtonForwardNext: build('actionButtonForwardNext', (g) => [darken(triangle(g, 1))]),
	actionButtonBackPrevious: build('actionButtonBackPrevious', (g) => [darken(triangle(g, -1))]),
	actionButtonBeginning: build('actionButtonBeginning', (g) => [
		darken(
			g.poly([
				[-D, -D],
				[-25.3, -D],
				[-25.3, D],
				[-D, D],
			]),
		),
		darken(
			g.poly([
				[D, -D],
				[-15, 0],
				[D, D],
			]),
		),
	]),
	actionButtonEnd: build('actionButtonEnd', (g) => [
		darken(
			g.poly([
				[25.3, -D],
				[D, -D],
				[D, D],
				[25.3, D],
			]),
		),
		darken(
			g.poly([
				[-D, -D],
				[15, 0],
				[-D, D],
			]),
		),
	]),
	actionButtonReturn: build('actionButtonReturn', (g) => {
		const { x, y, len } = { x: g.x.bind(g), y: g.y.bind(g), len: g.len.bind(g) };
		const ro = len(22);
		const ri = len(7);
		return [
			darken(
				`M ${x(-33.75)} ${y(-15.7)} L ${x(-33.75)} ${y(11.75)} A ${ro} ${ro} 10800000 -5400000 ` +
					`L ${x(2.7)} ${y(D)} A ${ro} ${ro} 5400000 -5400000 L ${x(24.7)} ${y(-16.2)} ` +
					`L ${x(34.5)} ${y(-16.2)} L ${x(16.85)} ${y(-D)} L ${x(-0.8)} ${y(-16.2)} ` +
					`L ${x(9)} ${y(-16.2)} L ${x(9)} ${y(11)} A ${ri} ${ri} 0 5400000 ` +
					`L ${x(-10.5)} ${y(18)} A ${ri} ${ri} 5400000 5400000 L ${x(-17.5)} ${y(-15.7)} Z`,
			),
		];
	}),
	actionButtonHome: build('actionButtonHome', (g) => [
		darkenLess(
			g.poly([
				[-24.75, 0],
				[24.75, 0],
				[24.75, D],
				[4.5, D],
				[4.5, 17.6],
				[-4.3, 17.6],
				[-4.3, D],
				[-24.75, D],
			]),
		),
		darkenLess(
			g.poly([
				[12.2, -30],
				[22, -30],
				[22, -10.5],
				[12.2, -20.5],
			]),
		),
		darken(
			g.poly([
				[0, -33],
				[32.25, 0],
				[-32.25, 0],
			]),
		),
		darken(
			g.poly([
				[-3.7, 17.75],
				[3.7, 17.75],
				[3.7, D],
				[-3.7, D],
			]),
		),
	]),
	actionButtonHelp: build('actionButtonHelp', (g) => [
		darken(g.poly(QUESTION)),
		darken(g.circle(-0.2, 27.4, 7)),
	]),
	actionButtonInformation: build('actionButtonInformation', (g) => [
		darken(g.circle(0, 0.4, 34)),
		lighten(g.circle(0, -24.7, 6)),
		lighten(
			g.poly([
				[-12, -11.2],
				[6, -11.2],
				[6, 22.5],
				[12, 22.5],
				[12, 25.5],
				[-12, 25.5],
				[-12, 22.5],
				[-6, 22.5],
				[-6, -8.2],
				[-12, -8.2],
			]),
		),
	]),
	actionButtonDocument: build('actionButtonDocument', (g) => [
		darkenLess(
			g.poly([
				[-24.75, -D],
				[9.25, -D],
				[25, -18],
				[25, D],
				[-24.75, D],
			]),
		),
		darken(
			g.poly([
				[9.25, -D],
				[25, -18],
				[9.25, -18],
			]),
		),
	]),
	actionButtonSound: build('actionButtonSound', (g) => [
		darken(
			g.poly([
				[-D, -12.4],
				[-13, -12.4],
				[8.4, -D],
				[8.4, D],
				[-13, 12.4],
				[-D, 12.4],
			]),
		),
		p(
			g.poly(
				[
					[17.5, -12.75],
					[34.5, -25.25],
				],
				false,
			) +
				' ' +
				g.poly(
					[
						[17.5, 0],
						[34.5, 0],
					],
					false,
				) +
				' ' +
				g.poly(
					[
						[17.5, 12.75],
						[34.5, 25.25],
					],
					false,
				),
			{ fill: 'none' },
		),
	]),
	actionButtonMovie: build('actionButtonMovie', (g) => [darken(g.poly(MOVIE))]),
};

export const EXACT_ACTION_BUTTON_PRESET_DEFINITIONS: Record<string, PresetShapeGeometryDefinition> =
	buttons;
