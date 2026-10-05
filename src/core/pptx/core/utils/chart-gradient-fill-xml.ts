/**
 * Pure `a:gradFill` building and comparison for chart gradients, split out of
 * `chart-gradient-fill-writer.ts` (which reconciles the result into a
 * `c:spPr`) to keep both within the repo's file-size limit.
 *
 * @module utils/chart-gradient-fill-xml
 */
import { OOXML_ANGLE_UNITS_PER_DEGREE, OOXML_PERCENT_UNITS } from '../constants';
import type { PptxChartGradientFill, XmlObject } from '../types';
import type { ResolveChartColor } from './chart-color-choice';
import { chartColorHex } from './chart-color-choice';
import { buildSrgbColorChoice, colorsEqual } from './color-xml-preservation';

type GetLocalName = (key: string) => string;
type GradientStop = PptxChartGradientFill['stops'][number];

/** `a:lin/@ang` default the parser assumes when none is authored (top to bottom). */
export const DEFAULT_CHART_GRADIENT_ANGLE = 90;
const CENTRE = 0.5;
const FULL_TURN = 360;
const PERCENT = 100;

function localNameOf(key: string): string {
	return key.replace(/^.*:/u, '');
}

function findKey(obj: XmlObject, local: string, getLocalName: GetLocalName): string | undefined {
	return Object.keys(obj).find((k) => getLocalName(k) === local);
}

function clamp(value: number, min: number, max: number): number {
	return Math.min(max, Math.max(min, value));
}

function percentUnits(fraction: number): number {
	return Math.round(clamp(fraction, 0, 1) * OOXML_PERCENT_UNITS);
}

function positionUnits(position: number): number {
	return percentUnits(position / PERCENT);
}

function opacityUnits(opacity: number | undefined): number {
	return opacity === undefined ? OOXML_PERCENT_UNITS : percentUnits(opacity);
}

function angleUnits(angle: number | undefined): number {
	const degrees = (((angle ?? DEFAULT_CHART_GRADIENT_ANGLE) % FULL_TURN) + FULL_TURN) % FULL_TURN;
	return (
		Math.round(degrees * OOXML_ANGLE_UNITS_PER_DEGREE) % (FULL_TURN * OOXML_ANGLE_UNITS_PER_DEGREE)
	);
}

function focalUnits(gradient: PptxChartGradientFill): [number, number] {
	const focal = gradient.focalPoint ?? { x: CENTRE, y: CENTRE };
	return [percentUnits(focal.x), percentUnits(focal.y)];
}

function sortedStops(gradient: PptxChartGradientFill): GradientStop[] {
	return [...gradient.stops].sort((a, b) => a.position - b.position);
}

/**
 * Whether two gradients serialize identically: compared at OOXML resolution
 * (`a:gs/@pos`, `a:alpha`, `a:lin/@ang`, `a:fillToRect`), colours
 * case-insensitively, an absent angle as {@link DEFAULT_CHART_GRADIENT_ANGLE}
 * and an absent focal point as the centre.
 */
export function chartGradientsEqual(a: PptxChartGradientFill, b: PptxChartGradientFill): boolean {
	if (a.type !== b.type || a.stops.length !== b.stops.length) {
		return false;
	}
	const left = sortedStops(a);
	const right = sortedStops(b);
	const stopsEqual = left.every(
		(stop, i) =>
			positionUnits(stop.position) === positionUnits(right[i].position) &&
			colorsEqual(stop.color, right[i].color) &&
			opacityUnits(stop.opacity) === opacityUnits(right[i].opacity),
	);
	if (!stopsEqual) {
		return false;
	}
	if (a.type === 'radial') {
		const [ax, ay] = focalUnits(a);
		const [bx, by] = focalUnits(b);
		return ax === bx && ay === by;
	}
	return angleUnits(a.angle) === angleUnits(b.angle);
}

/** The authored `a:gradFill` an edit is layered onto, with its parsed form. */
export interface AuthoredChartGradient {
	node: XmlObject;
	parsed: PptxChartGradientFill;
}

function colorChoiceOf(gs: XmlObject): XmlObject {
	const choice: XmlObject = {};
	for (const [key, value] of Object.entries(gs)) {
		if (!key.startsWith('@_')) {
			choice[key] = value;
		}
	}
	return choice;
}

function buildStop(
	stop: GradientStop,
	authored: AuthoredChartGradient | undefined,
	getLocalName: GetLocalName,
	resolveColor: ResolveChartColor | undefined,
): XmlObject {
	const pos = positionUnits(stop.position);
	const gsLst = authored
		? authored.node[findKey(authored.node, 'gsLst', getLocalName) ?? '']
		: undefined;
	const gsKey = gsLst ? findKey(gsLst as XmlObject, 'gs', getLocalName) : undefined;
	const rawStops = gsKey ? [(gsLst as XmlObject)[gsKey]].flat() : [];
	const authoredStop = (rawStops as XmlObject[]).find(
		(gs) => Number.parseInt(String(gs['@_pos'] ?? '0'), 10) === pos,
	);
	const parsedStop = authored?.parsed.stops.find((s) => positionUnits(s.position) === pos);
	if (
		authoredStop &&
		parsedStop &&
		resolveColor &&
		colorsEqual(resolveColor(authoredStop), stop.color) &&
		opacityUnits(parsedStop.opacity) === opacityUnits(stop.opacity)
	) {
		return { '@_pos': String(pos), ...colorChoiceOf(authoredStop) };
	}
	return { '@_pos': String(pos), ...buildSrgbColorChoice(chartColorHex(stop.color), stop.opacity) };
}

function buildPath(
	gradient: PptxChartGradientFill,
	authored: AuthoredChartGradient | undefined,
	getLocalName: GetLocalName,
): XmlObject {
	const authoredPath =
		authored?.parsed.type === 'radial'
			? (authored.node[findKey(authored.node, 'path', getLocalName) ?? ''] as XmlObject | undefined)
			: undefined;
	const path: XmlObject = { '@_path': String(authoredPath?.['@_path'] ?? 'circle') };
	const [x, y] = focalUnits(gradient);
	const authoredFocal =
		authored?.parsed.type === 'radial' ? focalUnits(authored.parsed) : undefined;
	if (authoredPath && authoredFocal && authoredFocal[0] === x && authoredFocal[1] === y) {
		// Same focus: keep the authored `a:fillToRect` (or its absence) verbatim.
		const rectKey = findKey(authoredPath, 'fillToRect', getLocalName);
		if (rectKey) {
			path[rectKey] = authoredPath[rectKey];
		}
		return path;
	}
	path['a:fillToRect'] = {
		'@_l': String(x),
		'@_t': String(y),
		'@_r': String(OOXML_PERCENT_UNITS - x),
		'@_b': String(OOXML_PERCENT_UNITS - y),
	};
	return path;
}

/**
 * Build an `a:gradFill` node (the element's content, keyed by the caller).
 * With `authored`, its attributes (`rotWithShape`, `flip`), `a:lin/@scaled`,
 * radial `a:path/@path` and `a:tileRect` carry over, and each stop whose
 * position, colour and opacity are unchanged re-emits its authored colour
 * choice. Without it, a fresh node: `rotWithShape="1"`, `a:lin scaled="0"`.
 */
export function buildChartGradFillXml(
	gradient: PptxChartGradientFill,
	authored?: AuthoredChartGradient,
	getLocalName: GetLocalName = localNameOf,
	resolveColor?: ResolveChartColor,
): XmlObject {
	const node: XmlObject = {};
	if (authored) {
		for (const [key, value] of Object.entries(authored.node)) {
			if (key.startsWith('@_')) {
				node[key] = value;
			}
		}
	} else {
		node['@_rotWithShape'] = '1';
	}
	node['a:gsLst'] = {
		'a:gs': sortedStops(gradient).map((stop) =>
			buildStop(stop, authored, getLocalName, resolveColor),
		),
	};
	if (gradient.type === 'radial') {
		node['a:path'] = buildPath(gradient, authored, getLocalName);
	} else {
		const authoredLin =
			authored?.parsed.type === 'linear'
				? (authored.node[findKey(authored.node, 'lin', getLocalName) ?? ''] as
						| XmlObject
						| undefined)
				: undefined;
		node['a:lin'] = {
			'@_ang': String(angleUnits(gradient.angle)),
			'@_scaled': String(authoredLin?.['@_scaled'] ?? '0'),
		};
	}
	const tileRectKey = authored ? findKey(authored.node, 'tileRect', getLocalName) : undefined;
	if (authored && tileRectKey) {
		node[tileRectKey] = authored.node[tileRectKey];
	}
	return node;
}
