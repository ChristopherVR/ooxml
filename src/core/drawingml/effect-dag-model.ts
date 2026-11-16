// The DrawingML effect DAG model (`a:effectDag` and `a:cont`, CT_EffectContainer, ECMA-376
// Part 1, 20.1.8.25 and 20.1.8.20) and its renderer-facing interpretation. Units are converted
// to plain numbers: fractions for percentages (1 = 100%), degrees for angles, EMU for lengths.
// `source` is the node a value was read from, so a host can resolve colours against its theme.
import type { DrawingEffectFields } from './effect-read';
import type { DrawingColor } from './types';

export interface DrawingEffectContainer<N = unknown> {
	kind: 'container';
	/** `sib`: every child reads the container's input; `tree`: each child reads the previous output. */
	type: 'sib' | 'tree';
	name?: string;
	children: DrawingDagNode<N>[];
}

/** A DAG colour or alpha operation (the EG_Effect members that are not list effects). */
export type DrawingDagColorOp<N = unknown> = { source?: N } & (
	| { kind: 'grayscale' }
	| { kind: 'biLevel'; threshold: number }
	| { kind: 'luminance'; brightness: number; contrast: number }
	| { kind: 'hsl'; hue: number; saturation: number; luminance: number }
	| { kind: 'tint'; hue: number; amount: number }
	| { kind: 'duotone'; colors: DrawingColor[] }
	| { kind: 'colorChange'; from?: DrawingColor; to?: DrawingColor; useAlpha: boolean }
	| { kind: 'colorReplace'; color?: DrawingColor }
	| { kind: 'alphaBiLevel'; threshold: number }
	| { kind: 'alphaCeiling' }
	| { kind: 'alphaFloor' }
	| { kind: 'alphaInverse'; color?: DrawingColor }
	| { kind: 'alphaOutset'; radiusEmu: number }
);

export type DrawingDagNode<N = unknown> =
	| DrawingEffectContainer<N>
	| { kind: 'effect'; effect: DrawingEffectFields; source?: N }
	| { kind: 'reference'; ref: string }
	| { kind: 'alphaMod'; container?: DrawingEffectContainer<N> }
	| { kind: 'blend'; blend: string; container?: DrawingEffectContainer<N> }
	| { kind: 'alphaModFix'; amount: number }
	| { kind: 'alphaReplace'; alpha: number }
	| { kind: 'colorOp'; op: DrawingDagColorOp<N> }
	| { kind: 'fill'; source?: N }
	| { kind: 'relativeOffset'; tx: number; ty: number }
	| {
			kind: 'transform';
			scaleX: number;
			scaleY: number;
			skewX: number;
			skewY: number;
			offsetXEmu: number;
			offsetYEmu: number;
	  }
	| { kind: 'unsupported'; name: string };

/**
 * What a painter applies for a DAG: an approximation of the graph as one effect list plus colour
 * operations on the shape. `approximations` lists, once each, what was simplified or dropped.
 */
export interface DrawingEffectDagPaint<N = unknown> {
	/** List-compatible effects on the paint path, in document order, with their nodes. */
	effects: { effect: DrawingEffectFields; source?: N }[];
	/** Colour and alpha operations applied to the shape's pixels, in order. */
	colorOps: DrawingDagColorOp<N>[];
	/** The product of `alphaModFix` amounts, or the last `alphaRepl` alpha, on the paint path. */
	opacity: number;
	/** The first relative offset (fractions of the shape box). */
	offset?: { tx: number; ty: number };
	/** The first transform effect. */
	transform?: Extract<DrawingDagNode, { kind: 'transform' }>;
	/** The first `a:blend` mode whose container is painted. */
	blend?: string;
	approximations: string[];
}
