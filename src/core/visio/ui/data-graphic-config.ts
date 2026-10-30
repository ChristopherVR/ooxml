import type { VisioPage } from '../model';
import type { VisioEdit } from '../edit-commands';
import { visioNextShapeId } from './shape-id';
import { visioDataGraphicMarker as marker, type VisioDataGraphicKind } from './data-graphic-rules';

/** Where a data graphic sits around its shape, as the nine-cell Position gallery lists them. */
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
/** The position each kind takes when none was chosen, matching where they were first drawn. */
export const VISIO_DATA_GRAPHIC_DEFAULT_POSITION: Record<
	Exclude<VisioDataGraphicKind, 'color'>,
	VisioDataGraphicPosition
> = { text: 'bottom', bar: 'bottom', icon: 'top-right' };

/** What a generated graphic remembers so Position, Configure and Auto Space can redo it. */
export interface VisioDataGraphicConfig {
	kind: Exclude<VisioDataGraphicKind, 'color'>;
	field: string;
	position?: VisioDataGraphicPosition;
	min?: number;
	max?: number;
}

const GAP = 0.05;
/** Centre of a graphic part around its owner; `slot` stacks several graphics at one position. */
export function visioDataGraphicCenter(
	owner: { x: number; y: number; width: number; height: number },
	part: { width: number; height: number },
	position: VisioDataGraphicPosition,
	slot: number,
): { x: number; y: number } {
	const ax = position.endsWith('left') ? -1 : position.endsWith('right') ? 1 : 0;
	const ay = position.startsWith('top') ? 1 : position.startsWith('bottom') ? -1 : 0;
	const step = 0.3 + GAP;
	return {
		x: owner.x + ax * (owner.width / 2 + part.width / 2 + GAP),
		// Above and below stack away from the shape; beside it they stack downward.
		y:
			owner.y +
			ay * (owner.height / 2 + part.height / 2 + GAP + slot * step) -
			(ay === 0 ? slot * step : 0),
	};
}

/** The marker text a graphic part stores: position and the resolved range. */
export function encodeDataGraphicConfig(config: VisioDataGraphicConfig): string {
	return [
		config.position ? `position=${config.position}` : '',
		config.min === undefined ? '' : `min=${config.min}`,
		config.max === undefined ? '' : `max=${config.max}`,
	]
		.filter(Boolean)
		.join(';');
}
function decode(
	kind: string,
	fieldName: string,
	format: string,
): VisioDataGraphicConfig | undefined {
	if (kind !== 'text' && kind !== 'bar' && kind !== 'icon') return undefined;
	if (!fieldName) return undefined;
	const values = new Map(format.split(';').map((item) => item.split('=') as [string, string]));
	const position = values.get('position') as VisioDataGraphicPosition | undefined;
	const number = (name: string) => {
		const text = values.get(name);
		return text === undefined || text === '' || !Number.isFinite(Number(text))
			? undefined
			: Number(text);
	};
	const [min, max] = [number('min'), number('max')];
	return {
		kind,
		field: fieldName,
		...(position && VISIO_DATA_GRAPHIC_POSITIONS.includes(position) ? { position } : {}),
		...(min === undefined ? {} : { min }),
		...(max === undefined ? {} : { max }),
	};
}

/** The graphics (text, bar, icon) of one owner in the order they were added; one per kind. */
export function visioDataGraphicConfigs(
	page: VisioPage,
	ownerId: string,
): VisioDataGraphicConfig[] {
	const result = new Map<string, VisioDataGraphicConfig>();
	for (const shape of page.shapes) {
		const row = marker(shape);
		if (!row || String(row.value ?? row.rawValue ?? '') !== ownerId) continue;
		const config = decode(row.label ?? '', row.prompt ?? '', row.format ?? '');
		if (config && !result.has(config.kind)) result.set(config.kind, config);
	}
	return [...result.values()];
}

type Regenerate = (
	page: VisioPage,
	shapeIds: readonly string[],
	options: VisioDataGraphicConfig & { slot: number; idStart: number },
) => VisioEdit[];

const nextAfter = (edits: readonly VisioEdit[], start: number) =>
	edits.reduce((next, edit) => {
		const id = 'shapeId' in edit ? Number(edit.shapeId) : NaN;
		return Number.isFinite(id) ? Math.max(next, id + 1) : next;
	}, start);

/**
 * Redraw the graphics of these owners at a position, or at their own positions when none is given
 * (Auto Space). Graphics that share a position are stacked in the order they were added, so they
 * never overlap. `regenerate` is the data graphic builder.
 */
export function visioRespaceDataGraphicEdits(
	page: VisioPage,
	ownerIds: readonly string[],
	regenerate: Regenerate,
	position?: VisioDataGraphicPosition,
): VisioEdit[] {
	const edits: VisioEdit[] = [];
	let next = Number(visioNextShapeId(page));
	for (const owner of ownerIds) {
		if (!page.shapes.some((shape) => shape.id === owner && !marker(shape))) continue;
		const used = new Map<string, number>();
		for (const config of visioDataGraphicConfigs(page, owner)) {
			const at = position ?? config.position ?? VISIO_DATA_GRAPHIC_DEFAULT_POSITION[config.kind];
			const slot = used.get(at) ?? 0;
			used.set(at, slot + 1);
			const made = regenerate(page, [owner], { ...config, position: at, slot, idStart: next });
			edits.push(...made);
			next = nextAfter(made, next);
		}
	}
	return edits;
}
