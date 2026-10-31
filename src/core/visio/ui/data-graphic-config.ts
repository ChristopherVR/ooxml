import type { VisioPage } from '../model';
import type { VisioEdit } from '../edit-commands';
import { visioNextShapeId } from './shape-id';
import { visioDataGraphicMarker as marker, type VisioDataGraphicKind } from './data-graphic-rules';
import {
	visioNativeGraphicItems,
	visioNativeGraphicOwners,
	visioSetNativeGraphicEdit,
} from './data-graphic-native';

import {
	VISIO_DATA_GRAPHIC_DEFAULT_POSITION,
	VISIO_DATA_GRAPHIC_POSITIONS,
	visioDataGraphicCenter,
	type VisioDataGraphicPosition,
} from '../data-graphic-place';
export {
	VISIO_DATA_GRAPHIC_DEFAULT_POSITION,
	VISIO_DATA_GRAPHIC_POSITIONS,
	visioDataGraphicCenter,
	type VisioDataGraphicPosition,
};

/** What a generated graphic remembers so Position, Configure and Auto Space can redo it. */
export interface VisioDataGraphicConfig {
	kind: Exclude<VisioDataGraphicKind, 'color'>;
	field: string;
	position?: VisioDataGraphicPosition;
	min?: number;
	max?: number;
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
	const native = visioNativeGraphicItems(page, ownerId);
	if (native.length) return native;
	const result = new Map<string, VisioDataGraphicConfig>();
	for (const shape of page.shapes) {
		const row = marker(shape);
		if (!row || String(row.value ?? row.rawValue ?? '') !== ownerId) continue;
		const config = decode(row.label ?? '', row.prompt ?? '', row.format ?? '');
		const key = `${config?.kind}\u0000${config?.field}`;
		if (config && !result.has(key)) result.set(key, config);
	}
	return [...result.values()];
}

type Regenerate = (
	page: VisioPage,
	shapeIds: readonly string[],
	options: VisioDataGraphicConfig & { slot: number; idStart: number; replace: 'field' },
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
		const native = visioNativeGraphicItems(page, owner);
		if (native.length) {
			// The callouts stack themselves in order at each position.
			edits.push(
				visioSetNativeGraphicEdit(
					page,
					owner,
					native.map((item) => ({ ...item, ...(position ? { position } : {}) })),
				),
			);
			continue;
		}
		if (!page.shapes.some((shape) => shape.id === owner && !marker(shape))) continue;
		const used = new Map<string, number>();
		for (const config of visioDataGraphicConfigs(page, owner)) {
			const at = position ?? config.position ?? VISIO_DATA_GRAPHIC_DEFAULT_POSITION[config.kind];
			const slot = used.get(at) ?? 0;
			used.set(at, slot + 1);
			const made = regenerate(page, [owner], {
				...config,
				position: at,
				slot,
				idStart: next,
				replace: 'field',
			});
			edits.push(...made);
			next = nextAfter(made, next);
		}
	}
	return edits;
}

/** Every shape that has a data graphic of text callouts, bars or icon sets, native or drawn. */
export function visioDataGraphicOwners(page: VisioPage): string[] {
	const owners = new Set(visioNativeGraphicOwners(page));
	for (const shape of page.shapes) {
		const row = marker(shape);
		const owner = row && String(row.value ?? row.rawValue ?? '');
		if (owner && owner !== 'legend') owners.add(owner);
	}
	return [...owners];
}
