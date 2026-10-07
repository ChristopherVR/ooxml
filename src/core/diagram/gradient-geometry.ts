import type { DiagramFill } from './types';

type Gradient = Extract<DiagramFill, { kind: 'gradient' }>;
export type RectGradientDirection =
	| 'center'
	| 'top-left'
	| 'top-right'
	| 'bottom-left'
	| 'bottom-right';
/** Native Office rectangular directions, measured from saved Excel center/corner fills. */
export const RECT_GRADIENT_DIRECTIONS: ReadonlyArray<{
	id: RectGradientDirection;
	label: string;
	x: number;
	y: number;
}> = [
	{ id: 'center', label: 'From Center', x: 0.5, y: 0.5 },
	{ id: 'top-left', label: 'From Top Left Corner', x: 0, y: 0 },
	{ id: 'top-right', label: 'From Top Right Corner', x: 1, y: 0 },
	{ id: 'bottom-left', label: 'From Bottom Left Corner', x: 0, y: 1 },
	{ id: 'bottom-right', label: 'From Bottom Right Corner', x: 1, y: 1 },
];

export function rectGradientFocus(id: RectGradientDirection): NonNullable<Gradient['fillToRect']> {
	const direction = RECT_GRADIENT_DIRECTIONS.find((direction) => direction.id === id);
	if (!direction) throw new RangeError(`Unknown rectangular gradient direction ${id}`);
	return { l: direction.x, t: direction.y, r: 1 - direction.x, b: 1 - direction.y };
}

export function rectGradientDirection(
	focus: Gradient['fillToRect'],
): RectGradientDirection | undefined {
	return RECT_GRADIENT_DIRECTIONS.find(
		({ x, y }) => focus?.l === x && focus.t === y && focus.r === 1 - x && focus.b === 1 - y,
	)?.id;
}

/** Replace geometry without changing stops, root flags or carried extensions. */
export function withDrawingGradientGeometry(
	fill: Gradient,
	type: 'linear' | 'rect',
	direction: RectGradientDirection = 'center',
): Gradient {
	const next = structuredClone(fill);
	next.tileRect = { l: 0, t: 0, r: 0, b: 0 };
	if (type === 'linear') {
		delete next.path;
		delete next.fillToRect;
		next.angle ??= 90;
		next.scaled = true;
	} else if (type === 'rect') {
		next.path = 'rect';
		next.fillToRect = rectGradientFocus(direction);
		delete next.angle;
		delete next.scaled;
		if (direction !== 'center') {
			const { l, t } = next.fillToRect;
			next.tileRect = {
				l: l === 0 ? -1 : 0,
				t: t === 0 ? -1 : 0,
				r: l === 1 ? -1 : 0,
				b: t === 1 ? -1 : 0,
			};
		}
	} else throw new RangeError(`Unknown gradient type ${type}`);
	return next;
}
