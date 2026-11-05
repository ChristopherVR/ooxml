/**
 * Top/bottom curve definitions for the `prstTxWarp` presets whose two paths
 * are plain functions of x but were drawn from a single reconstructed
 * baseline before: the shape presets (stop, triangle, chevron), the curves,
 * the waves, and the fade, slant and cascade families.
 *
 * Every number here was measured against PowerPoint 16: each preset was
 * exported with a row of full-block glyphs, whose ink fills exactly the area
 * between the two paths, and the boundary was read off per column (at the
 * default adjust and at two or more others). The guide shapes follow the
 * `presetTextWarpDefinitions.xml` structure the existing definitions were
 * transcribed from; only the constants the measurement fixed are used.
 *
 * Coordinates are guide names resolved against the normalised 100000 box by
 * `text-warp-preset-sampler.ts`. `adj` is the first adjust (`adj1` in the
 * wave presets). The wave presets' second adjust (`adj2`, a horizontal
 * shift) is not modelled yet and reads as its default of 0.
 */
import type {
	WarpChainPart,
	WarpCurveSegment,
	WarpPresetDefinition,
} from './text-warp-preset-definitions';

type Piece = WarpChainPart['seg'];

const line = (startY: string, endY: string): Piece => ({ type: 'line', startY, endY });
const cubic = (startY: string, ctrl1Y: string, ctrl2Y: string, endY: string): Piece => ({
	type: 'cubic',
	startY,
	ctrl1Y,
	ctrl2Y,
	endY,
});
const quad = (startY: string, ctrlY: string, endY: string): Piece => ({
	type: 'quad',
	startY,
	ctrlY,
	endY,
});
const part = (x0: string, x1: string, seg: Piece): WarpChainPart => ({ x0, x1, seg });
const chain = (...parts: WarpChainPart[]): WarpCurveSegment => ({ type: 'chain', parts });

/** Two straight pieces meeting at the horizontal centre. */
const peak = (leftY: string, midY: string, rightY: string): WarpCurveSegment =>
	chain(part('l', 'hc', line(leftY, midY)), part('hc', 'r', line(midY, rightY)));

/** Three pieces: a ramp up to `x1`, a plateau, a ramp back from `x2`. */
const trapezoid = (edgeY: string, plateauY: string): WarpCurveSegment =>
	chain(
		part('l', 'x1', line(edgeY, plateauY)),
		part('x1', 'x2', line(plateauY, plateauY)),
		part('x2', 'r', line(plateauY, edgeY)),
	);

/** The wave guides shared by all four wave presets (`adj` is `adj1`). */
const waveGuides = (maxAdj: number) =>
	[
		{ name: 'a1', formula: `pin 0 adj ${maxAdj}` },
		{ name: 'y1', formula: '*/ h a1 100000' },
		{ name: 'dy2', formula: '*/ y1 10 3' },
		{ name: 'y2', formula: '+- y1 0 dy2' },
		{ name: 'y3', formula: '+- y1 dy2 0' },
		{ name: 'y4', formula: '+- b 0 y1' },
		{ name: 'y5', formula: '+- y4 0 dy2' },
		{ name: 'y6', formula: '+- y4 dy2 0' },
	] as const;

/** `adj2` slides the two paths against each other: one is shortened at the right, the other at the left. */
const shiftGuides = [
	{ name: 'a2', formula: 'pin -10000 adj2 10000' },
	{ name: 'dxa', formula: '*/ w a2 50000' },
	{ name: 'pos', formula: '?: dxa dxa 0' },
	{ name: 'neg', formula: '+- pos 0 dxa' },
	{ name: 'tsx', formula: '+- l neg 0' },
	{ name: 'tex', formula: '+- r 0 pos' },
	{ name: 'bsx', formula: '+- l pos 0' },
	{ name: 'bex', formula: '+- r 0 neg' },
] as const;

const shiftExtents = {
	topX: { startX: 'tsx', endX: 'tex' },
	bottomX: { startX: 'bsx', endX: 'bex' },
} as const;

const halves = (a: string, b: string, c: string, d: string): WarpCurveSegment =>
	chain(part('l', 'hc', cubic(a, b, c, d)), part('hc', 'r', cubic(a, b, c, d)));

const sideGuides = [
	{ name: 'dy', formula: '*/ a h 100000' },
	{ name: 'yb', formula: '+- b 0 dy' },
] as const;

/** `x1`/`x2`: where a ramp reaches its plateau and where it leaves it. */
const rampGuides = [
	{ name: 'x1', formula: '*/ w a 100000' },
	{ name: 'x2', formula: '+- r 0 x1' },
] as const;

const curveGuides = [
	{ name: 'a', formula: 'pin 0 adj 56000' },
	{ name: 'dy1', formula: '*/ h a 100000' },
	{ name: 'dy2', formula: '*/ dy1 3 4' },
	{ name: 'dy3', formula: '*/ dy1 5 4' },
	{ name: 'dy5', formula: '*/ dy1 3 8' },
	{ name: 'yb0', formula: '+- b 0 dy5' },
	{ name: 'yb1', formula: '+- b dy5 0' },
] as const;

const cascadeGuides = [
	{ name: 'a', formula: 'pin 28572 adj 100000' },
	{ name: 'dy', formula: '*/ a h 100000' },
	{ name: 'dy2', formula: '+- b 0 dy' },
	{ name: 'yq', formula: '*/ dy2 1 4' },
] as const;

export const WARP_PRESET_DEFINITIONS_EXTRA: Readonly<Record<string, WarpPresetDefinition>> = {
	textStop: {
		defaultAdj: 25000,
		gdLst: [
			{ name: 'a', formula: 'pin 0 adj 50000' },
			...sideGuides,
			{ name: 'x1', formula: '+- wd3 0 0' },
			{ name: 'x2', formula: '+- r 0 wd3' },
		],
		top: trapezoid('dy', 't'),
		bottom: trapezoid('yb', 'b'),
	},
	textTriangle: {
		defaultAdj: 50000,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 100000' }, ...sideGuides],
		top: peak('dy', 't', 'dy'),
		bottom: peak('b', 'b', 'b'),
	},
	textTriangleInverted: {
		defaultAdj: 50000,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 100000' }, ...sideGuides],
		top: peak('t', 't', 't'),
		bottom: peak('dy', 'b', 'dy'),
	},
	textChevron: {
		defaultAdj: 25000,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 100000' }, ...sideGuides],
		top: peak('dy', 't', 'dy'),
		bottom: peak('b', 'yb', 'b'),
	},
	textChevronInverted: {
		defaultAdj: 75000,
		gdLst: [{ name: 'a', formula: 'pin 50000 adj 100000' }, ...sideGuides],
		top: peak('t', 'yb', 't'),
		bottom: peak('dy', 'b', 'dy'),
	},
	textCurveUp: {
		defaultAdj: 45977,
		gdLst: curveGuides,
		top: cubic('dy1', 'dy3', 'dy2', 't'),
		bottom: quad('yb0', 'yb1', 'yb0'),
	},
	textCurveDown: {
		defaultAdj: 45977,
		gdLst: curveGuides,
		top: cubic('t', 'dy2', 'dy3', 'dy1'),
		bottom: quad('yb0', 'yb1', 'yb0'),
	},
	textFadeRight: {
		defaultAdj: 33333,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 49000' }, ...sideGuides],
		top: line('t', 'dy'),
		bottom: line('b', 'yb'),
	},
	textFadeLeft: {
		defaultAdj: 33333,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 49000' }, ...sideGuides],
		top: line('dy', 't'),
		bottom: line('yb', 'b'),
	},
	textFadeUp: {
		defaultAdj: 33333,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 49000' }, ...rampGuides],
		// The top path is only the plateau (`x1..x2`); the bottom path is the full
		// width, so the block is pinched toward the top.
		top: line('t', 't'),
		bottom: line('b', 'b'),
		topX: { startX: 'x1', endX: 'x2' },
	},
	textFadeDown: {
		defaultAdj: 33333,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 49000' }, ...rampGuides],
		top: line('t', 't'),
		bottom: line('b', 'b'),
		bottomX: { startX: 'x1', endX: 'x2' },
	},
	textSlantUp: {
		defaultAdj: 55556,
		gdLst: [{ name: 'a', formula: 'pin 0 adj 71428' }, ...sideGuides],
		top: line('dy', 't'),
		bottom: line('b', 'yb'),
	},
	textSlantDown: {
		defaultAdj: 44444,
		gdLst: [{ name: 'a', formula: 'pin 28572 adj 100000' }, ...sideGuides],
		top: line('t', 'yb'),
		bottom: line('dy', 'b'),
	},
	textCascadeUp: {
		defaultAdj: 44444,
		gdLst: cascadeGuides,
		top: line('yq', 't'),
		bottom: line('b', 'dy'),
	},
	textCascadeDown: {
		defaultAdj: 44444,
		gdLst: cascadeGuides,
		top: line('t', 'yq'),
		bottom: line('dy', 'b'),
	},
	textWave1: {
		defaultAdj: 12500,
		gdLst: [...waveGuides(20000), ...shiftGuides],
		...shiftExtents,
		top: cubic('y1', 'y2', 'y3', 'y1'),
		bottom: cubic('y4', 'y5', 'y6', 'y4'),
	},
	textWave2: {
		defaultAdj: 12500,
		gdLst: [...waveGuides(20000), ...shiftGuides],
		...shiftExtents,
		top: cubic('y1', 'y3', 'y2', 'y1'),
		bottom: cubic('y4', 'y6', 'y5', 'y4'),
	},
	textWave4: {
		defaultAdj: 6250,
		gdLst: [...waveGuides(12500), ...shiftGuides],
		...shiftExtents,
		top: halves('y1', 'y3', 'y2', 'y1'),
		bottom: halves('y4', 'y6', 'y5', 'y4'),
	},
	textDoubleWave1: {
		defaultAdj: 6250,
		gdLst: [...waveGuides(12500), ...shiftGuides],
		...shiftExtents,
		top: halves('y1', 'y2', 'y3', 'y1'),
		bottom: halves('y4', 'y5', 'y6', 'y4'),
	},
};
