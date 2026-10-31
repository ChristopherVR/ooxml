/** Where a data graphic sits around its shape, as the Position gallery lists them. */
export type VisioDataGraphicPosition =
	| 'top-left'
	| 'top'
	| 'top-right'
	| 'left'
	| 'right'
	| 'bottom-left'
	| 'bottom'
	| 'bottom-right';
export const VISIO_DATA_GRAPHIC_POSITIONS: readonly VisioDataGraphicPosition[] = [
	'top-left',
	'top',
	'top-right',
	'left',
	'right',
	'bottom-left',
	'bottom',
	'bottom-right',
];
/** The callout kinds a graphic is made of (Color by Value changes the fill and has no callout). */
export type VisioDataGraphicCalloutKind = 'text' | 'bar' | 'icon';
/** The position each kind takes when none was chosen. */
export const VISIO_DATA_GRAPHIC_DEFAULT_POSITION: Record<
	VisioDataGraphicCalloutKind,
	VisioDataGraphicPosition
> = { text: 'bottom', bar: 'bottom', icon: 'top-right' };

/** Gap between a callout and its shape, and the pitch of stacked callouts, in inches. */
export const VISIO_DATA_GRAPHIC_GAP = 0.05;
export const VISIO_DATA_GRAPHIC_STACK_STEP = 0.3 + VISIO_DATA_GRAPHIC_GAP;

const SIDE: Record<VisioDataGraphicPosition, { ax: -1 | 0 | 1; ay: -1 | 0 | 1 }> = {
	'top-left': { ax: -1, ay: 1 },
	top: { ax: 0, ay: 1 },
	'top-right': { ax: 1, ay: 1 },
	left: { ax: -1, ay: 0 },
	right: { ax: 1, ay: 0 },
	'bottom-left': { ax: -1, ay: -1 },
	bottom: { ax: 0, ay: -1 },
	'bottom-right': { ax: 1, ay: -1 },
};

/** Centre of a graphic part around its owner; `slot` stacks several graphics at one position. */
export function visioDataGraphicCenter(
	owner: { x: number; y: number; width: number; height: number },
	part: { width: number; height: number },
	position: VisioDataGraphicPosition,
	slot: number,
): { x: number; y: number } {
	const { ax, ay } = SIDE[position];
	const step = VISIO_DATA_GRAPHIC_STACK_STEP;
	return {
		x: owner.x + ax * (owner.width / 2 + part.width / 2 + VISIO_DATA_GRAPHIC_GAP),
		// Above and below stack away from the shape; beside it they stack downward.
		y:
			owner.y +
			ay * (owner.height / 2 + part.height / 2 + VISIO_DATA_GRAPHIC_GAP + slot * step) -
			(ay === 0 ? slot * step : 0),
	};
}

/**
 * The `visDGDefaultPos` point a position is stored as: 2 is the centre of the shape's extent,
 * 5 is beyond its far edge and -1 beyond its near edge, on each axis, as Visio's own cell reads.
 */
export function visioDataGraphicPoint(position: VisioDataGraphicPosition): {
	x: number;
	y: number;
} {
	const { ax, ay } = SIDE[position];
	return { x: ax === 0 ? 2 : ax === 1 ? 5 : -1, y: ay === 0 ? 2 : ay === 1 ? 5 : -1 };
}

/** The position a stored point stands for, when it is one of the eight. */
export function visioDataGraphicPositionOf(
	x: number,
	y: number,
): VisioDataGraphicPosition | undefined {
	return VISIO_DATA_GRAPHIC_POSITIONS.find((position) => {
		const point = visioDataGraphicPoint(position);
		return point.x === x && point.y === y;
	});
}
