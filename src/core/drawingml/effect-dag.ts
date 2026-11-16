// The `a:effectDag` reader (CT_EffectContainer, ECMA-376 Part 1, 20.1.8.25) over a
// DrawingNodeView, and its interpretation as one paint (`DrawingEffectDagPaint`). Attribute
// defaults are the schema's; each EG_Effect member is typed, anything else stays `unsupported`.
import { parseDrawingFraction, parseDrawingHueDegrees } from '../color/color-primitives';
import type {
	DrawingDagColorOp,
	DrawingDagNode,
	DrawingEffectContainer,
	DrawingEffectDagPaint,
} from './effect-dag-model';
import { readDrawingEffect } from './effect-read';
import { drawingBoolean } from './gradient-read';
import type { DrawingNodeView } from './node-view';

export * from './effect-dag-model';

/** Reads `a:effectDag` or `a:cont`; `undefined` without a node. */
export function readDrawingEffectContainer<N>(
	view: DrawingNodeView<N>,
	node: N | undefined,
): DrawingEffectContainer<N> | undefined {
	if (node === undefined) return undefined;
	const type = view.attribute(node, 'type')?.trim();
	const name = view.attribute(node, 'name')?.trim();
	return {
		kind: 'container',
		type: type === 'tree' ? 'tree' : 'sib',
		...(name ? { name } : {}),
		children: view
			.elements(node)
			.map(({ name: child, node: element }) => readDagNode(view, child, element)),
	};
}

function readDagNode<N>(view: DrawingNodeView<N>, name: string, node: N): DrawingDagNode<N> {
	const fraction = (attribute: string, fallback: number) =>
		parseDrawingFraction(view.attribute(node, attribute)) ?? fallback;
	const degrees = (attribute: string) =>
		parseDrawingHueDegrees(view.attribute(node, attribute)) ?? 0;
	const emu = (attribute: string) => {
		const value = Number.parseInt(view.attribute(node, attribute) ?? '', 10);
		return Number.isFinite(value) ? value : 0;
	};
	const op = (value: DrawingDagColorOp<N>): DrawingDagNode<N> => ({
		kind: 'colorOp',
		op: { ...value, source: node },
	});
	const colorIn = (child: string) => {
		const holder = view.child(node, child);
		return holder === undefined ? undefined : view.color(holder);
	};
	const effect = readDrawingEffect(view, name, node);
	if (effect) return { kind: 'effect', effect, source: node };
	switch (name) {
		case 'cont':
			return readDrawingEffectContainer(view, node)!;
		case 'effect':
			return { kind: 'reference', ref: view.attribute(node, 'ref') ?? '' };
		case 'alphaMod': {
			const container = readDrawingEffectContainer(view, view.child(node, 'cont'));
			return container ? { kind: 'alphaMod', container } : { kind: 'alphaMod' };
		}
		case 'blend': {
			const container = readDrawingEffectContainer(view, view.child(node, 'cont'));
			const blend = view.attribute(node, 'blend')?.trim() || 'over';
			return container ? { kind: 'blend', blend, container } : { kind: 'blend', blend };
		}
		case 'alphaModFix':
			return { kind: 'alphaModFix', amount: fraction('amt', 1) };
		case 'alphaRepl':
			return { kind: 'alphaReplace', alpha: fraction('a', 1) };
		case 'grayscl':
			return op({ kind: 'grayscale' });
		case 'biLevel':
			return op({ kind: 'biLevel', threshold: fraction('thresh', 0.5) });
		case 'lum':
			return op({
				kind: 'luminance',
				brightness: fraction('bright', 0),
				contrast: fraction('contrast', 0),
			});
		case 'hsl':
			return op({
				kind: 'hsl',
				hue: degrees('hue'),
				saturation: fraction('sat', 0),
				luminance: fraction('lum', 0),
			});
		case 'tint':
			return op({ kind: 'tint', hue: degrees('hue'), amount: fraction('amt', 0) });
		case 'duotone':
			return op({ kind: 'duotone', colors: view.colors(node) });
		case 'clrChange': {
			const from = colorIn('clrFrom');
			const to = colorIn('clrTo');
			return op({
				kind: 'colorChange',
				...(from ? { from } : {}),
				...(to ? { to } : {}),
				useAlpha: drawingBoolean(view.attribute(node, 'useA')) ?? true,
			});
		}
		case 'clrRepl': {
			const color = view.color(node);
			return op({ kind: 'colorReplace', ...(color ? { color } : {}) });
		}
		case 'alphaBiLevel':
			return op({ kind: 'alphaBiLevel', threshold: fraction('thresh', 0.5) });
		case 'alphaCeiling':
			return op({ kind: 'alphaCeiling' });
		case 'alphaFloor':
			return op({ kind: 'alphaFloor' });
		case 'alphaInv': {
			const color = view.color(node);
			return op({ kind: 'alphaInverse', ...(color ? { color } : {}) });
		}
		case 'alphaOutset':
			return op({ kind: 'alphaOutset', radiusEmu: emu('rad') });
		case 'fill':
			return { kind: 'fill', source: node };
		case 'relOff':
			return { kind: 'relativeOffset', tx: fraction('tx', 0), ty: fraction('ty', 0) };
		case 'xfrm':
		case 'xfrmEffect':
			return {
				kind: 'transform',
				scaleX: fraction('sx', 1),
				scaleY: fraction('sy', 1),
				skewX: degrees('kx'),
				skewY: degrees('ky'),
				offsetXEmu: emu('tx'),
				offsetYEmu: emu('ty'),
			};
		default:
			return { kind: 'unsupported', name };
	}
}

/**
 * Interprets a DAG as one paint. The shape is the DAG's input; along the paint path list
 * effects are collected, colour operations accumulate in order and alpha multiplies.
 * Approximations, each named once in `approximations`:
 * - a `sib` container with several children composites their outputs as one paint;
 * - `a:alphaMod` masks with its container's alpha: the container is not painted and the mask is
 *   reduced to the `alphaModFix`/`alphaRepl` alpha it yields;
 * - a non-`over` `a:blend` paints its container normally and reports the mode in `blend`;
 * - `a:effect` references, `a:fill` and unknown members are not applied;
 * - transforms and offsets are reported, the first of each, not applied to the effects.
 */
export function interpretDrawingEffectDag<N>(
	dag: DrawingEffectContainer<N>,
): DrawingEffectDagPaint<N> {
	const paint: DrawingEffectDagPaint<N> = {
		effects: [],
		colorOps: [],
		opacity: 1,
		approximations: [],
	};
	const note = (text: string) => {
		if (!paint.approximations.includes(text)) paint.approximations.push(text);
	};
	const containerAlpha = (container: DrawingEffectContainer<N>): number => {
		let alpha = 1;
		for (const child of container.children) {
			if (child.kind === 'alphaModFix') alpha *= child.amount;
			else if (child.kind === 'alphaReplace') alpha = child.alpha;
			else if (child.kind === 'container') alpha *= containerAlpha(child);
		}
		return alpha;
	};
	const visit = (container: DrawingEffectContainer<N>) => {
		if (container.type === 'sib' && container.children.length > 1)
			note('sibling outputs composited as one paint');
		for (const node of container.children) {
			switch (node.kind) {
				case 'container':
					visit(node);
					break;
				case 'effect':
					paint.effects.push(
						node.source === undefined
							? { effect: node.effect }
							: { effect: node.effect, source: node.source },
					);
					break;
				case 'colorOp':
					paint.colorOps.push(node.op);
					break;
				case 'alphaModFix':
					paint.opacity *= node.amount;
					break;
				case 'alphaReplace':
					paint.opacity = node.alpha;
					break;
				case 'alphaMod':
					note('alphaMod reduced to its container alpha');
					if (node.container) paint.opacity *= containerAlpha(node.container);
					break;
				case 'blend':
					if (node.blend !== 'over') note(`blend ${node.blend} painted as over`);
					paint.blend ??= node.blend;
					if (node.container) visit(node.container);
					break;
				case 'relativeOffset':
					note('relOff reported, not applied');
					paint.offset ??= { tx: node.tx, ty: node.ty };
					break;
				case 'transform':
					note('xfrm reported, not applied');
					paint.transform ??= node;
					break;
				case 'reference':
					note('effect references not applied');
					break;
				case 'fill':
					note('fill effect not applied');
					break;
				case 'unsupported':
					note(`unsupported ${node.name}`);
					break;
			}
		}
	};
	visit(dag);
	paint.opacity = Math.min(1, Math.max(0, paint.opacity));
	return paint;
}
