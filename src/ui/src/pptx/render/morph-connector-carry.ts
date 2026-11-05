/**
 * An elbow or curved connector carried the way PowerPoint's Morph carries it.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object, 2 s,
 * a 6 pt `bentConnector3` and `curvedConnector3`):
 *
 *  - the stroke keeps its width for the whole morph (8 px on every frame while
 *    the box grows 3x), where our journey scales the wrapper and the stroke with it;
 *  - the route is re-evaluated per frame from the interpolated box: the bend
 *    stays at its `adj` fraction of the box (0.50 on every frame), the
 *    segments grow with the box and the corners stay square;
 *  - a flip is a SIGNED extent that passes through zero: an elbow flipped
 *    horizontally narrows to a vertical bar and opens out mirrored, it does
 *    not swap sides or cross-fade;
 *  - an `adj` change moves the bend linearly from one fraction to the other.
 *
 * So each frame's path is the pair's unit-box route (evaluated unflipped, the
 * two ends' paths blended when their `adj` differ) mapped onto the interpolated
 * centre, SIGNED extent and rotation, drawn in slide coordinates. The strand
 * undoes the whole journey (`transform: W(p)^-1`), so the wrapper and the
 * inverse cancel and the stroke is drawn at its own width, uniformly, whatever
 * the anisotropic scale of the journey.
 *
 * Both halves of a pair get the same path, so a ghost and its incoming half
 * stay on top of each other. Arrowheads and dashes are not touched here.
 *
 * @module render/morph-connector-carry
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getConnectorPathGeometry, hasShapeProperties } from 'ooxml-core/pptx';

import { connectorNeedsPath, getCompoundLineOffsets } from './connector-style';
import type { Matrix } from './morph-line-carry';
import {
	centre,
	EPSILON,
	IDENTITY,
	invert,
	journeyOf,
	lineWidth,
	precise,
	STOPS,
	wrapperAt,
} from './morph-line-carry';
import { reachFactors } from './morph-line-reach';
import { shortestRotationTarget } from './morph-rotation';
import { parseSvgPath } from './morph-svg-path';
import { bakeStops, lerp, round, sameStructure } from './morph-shape-carry-path';
import type { MorphAnimationStyle, MorphPair, SvgPathCommand } from './morph-types';

/** Which half of a pair a carry is painted on: its box is the local frame. */
export type MorphConnectorHalf = 'from' | 'to';

/** The unit box the preset route is evaluated in (its numbers read as thousandths of the box). */
const UNIT = 1000;

/** Whether `element` is a single-strand elbow or curved connector the carry can re-path. */
export function isCarriedPathConnector(element: PptxElement): boolean {
	if (element.type !== 'connector' || !hasShapeProperties(element)) return false;
	return (
		connectorNeedsPath(element.shapeType) &&
		lineWidth(element) > 0 &&
		getCompoundLineOffsets(element.shapeStyle?.compoundLine, lineWidth(element)).length === 1
	);
}

/** The connector's unflipped route in the unit box, or `undefined` when it is not a plain preset. */
function unitRoute(element: PptxElement): SvgPathCommand[] | undefined {
	if (!hasShapeProperties(element)) return undefined;
	const { pathData } = getConnectorPathGeometry({
		...element,
		width: UNIT,
		height: UNIT,
		flipHorizontal: false,
		flipVertical: false,
	});
	const commands = parseSvgPath(pathData);
	const known = commands.every((command) => 'MLQCZ'.includes(command.type));
	return known && commands.length > 1 ? commands : undefined;
}

const sign = (flipped: boolean | undefined): number => (flipped ? -1 : 1);

/** The route of the pair at eased progress `p`, in slide coordinates. */
function routeAt(
	a: readonly SvgPathCommand[],
	b: readonly SvgPathCommand[],
	from: PptxElement,
	to: PptxElement,
	p: number,
): string {
	const [fcx, fcy] = centre(from);
	const [tcx, tcy] = centre(to);
	const cx = lerp(fcx, tcx, p);
	const cy = lerp(fcy, tcy, p);
	const [reachX, reachY] = reachFactors(from, to, p);
	const sw =
		lerp(sign(from.flipHorizontal) * from.width, sign(to.flipHorizontal) * to.width, p) * reachX;
	const sh =
		lerp(sign(from.flipVertical) * from.height, sign(to.flipVertical) * to.height, p) * reachY;
	const fromRot = from.rotation ?? 0;
	const turn = lerp(fromRot, shortestRotationTarget(fromRot, to.rotation ?? 0), p);
	const [cos, sin] = [Math.cos((turn * Math.PI) / 180), Math.sin((turn * Math.PI) / 180)];
	const at = (x: number, y: number): [number, number] => {
		const lx = (x / UNIT - 0.5) * sw;
		const ly = (y / UNIT - 0.5) * sh;
		return [cx + cos * lx - sin * ly, cy + sin * lx + cos * ly];
	};
	return a
		.map((command, index) => {
			const other = b[index]?.values ?? command.values;
			const values: number[] = [];
			for (let k = 0; k + 1 < command.values.length; k += 2) {
				const x = lerp(command.values[k] ?? 0, other[k] ?? 0, p);
				const y = lerp(command.values[k + 1] ?? 0, other[k + 1] ?? 0, p);
				values.push(...at(x, y).map(round));
			}
			return `${command.type}${values.join(' ')}`;
		})
		.join(' ');
}

/**
 * The baked keyframes body for the painted strand of one half of a pair of
 * elbow or curved connectors whose box, flips, rotation or `adj` differ, or
 * `undefined` when the pair is not one.
 */
export function buildMorphConnectorCarry(
	from: PptxElement,
	to: PptxElement,
	half: MorphConnectorHalf,
): string | undefined {
	if (!isCarriedPathConnector(from) || !isCarriedPathConnector(to)) return undefined;
	if (!hasShapeProperties(from) || !hasShapeProperties(to)) return undefined;
	if ((from.shapeType ?? '').toLowerCase() !== (to.shapeType ?? '').toLowerCase()) return undefined;
	const a = unitRoute(from);
	const b = unitRoute(to);
	if (!a || !b || !sameStructure(a, b)) return undefined;
	const sameRoute = a.every((command, index) =>
		command.values.every((value, k) => Math.abs(value - (b[index]?.values[k] ?? value)) < 1e-6),
	);
	const unchanged =
		sameRoute &&
		Math.abs(from.width - to.width) <= EPSILON &&
		Math.abs(from.height - to.height) <= EPSILON &&
		(from.rotation ?? 0) === (to.rotation ?? 0) &&
		Boolean(from.flipHorizontal) === Boolean(to.flipHorizontal) &&
		Boolean(from.flipVertical) === Boolean(to.flipVertical);
	if (unchanged) return undefined;
	const journey = journeyOf(from, to, half);
	const w0 = lineWidth(from);
	const w1 = lineWidth(to);
	const progress = Array.from({ length: STOPS + 1 }, (_, i) => i / STOPS);
	const stops = new Map<number, { matrix: Matrix; d: string }>();
	for (const p of progress) {
		const inverse = invert(wrapperAt(journey, p));
		if (!inverse) return undefined;
		stops.set(p, { matrix: inverse, d: routeAt(a, b, from, to, p) });
	}
	return bakeStops(progress, (p) => {
		const stop = stops.get(p) ?? { matrix: IDENTITY, d: '' };
		return [
			'transform-origin: 0px 0px',
			'vector-effect: none',
			`transform: matrix(${stop.matrix.map(precise).join(', ')})`,
			`d: path('${stop.d}')`,
			`stroke-width: ${round(lerp(w0, w1, p))}px`,
		];
	});
}

/** The animation of one half's painted strand (`target: 'line'`), or none. */
export function buildMorphConnectorAnimations(
	pair: MorphPair,
	id: string,
	half: MorphConnectorHalf,
	durationMs: number,
	tag: string,
): MorphAnimationStyle[] {
	const keyframes = buildMorphConnectorCarry(pair.fromElement, pair.toElement, half);
	if (!keyframes) return [];
	const name = `pptx-morph-connector-${tag}-${id.replace(/[^a-zA-Z0-9]/gu, '')}`;
	return [
		{
			elementId: id,
			animation: `${name} ${durationMs}ms linear forwards`,
			keyframes: `\n@keyframes ${name} {\n${keyframes}\n}`,
			target: 'line',
		},
	];
}
